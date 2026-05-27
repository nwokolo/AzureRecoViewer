import axios, { AxiosError } from "axios";

const ARM_BASE_URL = "https://management.azure.com";

export type ArmListResponse<T> = {
  value: T[];
  nextLink?: string;
};

function normalizeArmError(error: unknown): Error {
  if (axios.isAxiosError(error)) {
    const axiosError = error as AxiosError<{ error?: { message?: string; code?: string } }>;
    const message =
      axiosError.response?.data?.error?.message ??
      axiosError.response?.statusText ??
      axiosError.message;
    const code = axiosError.response?.data?.error?.code;
    return new Error(code ? `${code}: ${message}` : message);
  }

  if (error instanceof Error) {
    return error;
  }

  return new Error("Unknown ARM request failure");
}

async function withRetry<T>(operation: () => Promise<T>, attempt = 0): Promise<T> {
  try {
    return await operation();
  } catch (error) {
    if (!axios.isAxiosError(error) || attempt >= 3) {
      throw normalizeArmError(error);
    }

    const status = error.response?.status;
    if (status !== 429 && status !== 503) {
      throw normalizeArmError(error);
    }

    const retryAfterSecondsHeader =
      error.response?.headers["retry-after"] ??
      error.response?.headers["x-ms-ratelimit-microsoft.consumption-retry-after"];

    const retryAfterSeconds = Number.parseInt(String(retryAfterSecondsHeader ?? "1"), 10);
    const delayMs = Number.isFinite(retryAfterSeconds)
      ? Math.max(retryAfterSeconds, 1) * 1000
      : 1500;

    await new Promise((resolve) => setTimeout(resolve, delayMs));
    return withRetry(operation, attempt + 1);
  }
}

export async function armGet<T>(url: string, token: string): Promise<T> {
  return withRetry(async () => {
    const response = await axios.get<T>(url, {
      headers: {
        Authorization: `Bearer ${token}`,
      },
    });

    return response.data;
  });
}

export async function armListAll<T>(url: string, token: string): Promise<T[]> {
  const allItems: T[] = [];
  let nextUrl: string | undefined = url;

  while (nextUrl) {
    const currentUrl = nextUrl;
    const page: ArmListResponse<T> = await armGet<ArmListResponse<T>>(currentUrl, token);
    allItems.push(...(page.value ?? []));
    nextUrl = page.nextLink;
  }

  return allItems;
}

export function armUrl(pathOrAbsoluteUrl: string): string {
  if (pathOrAbsoluteUrl.startsWith("https://")) {
    return pathOrAbsoluteUrl;
  }

  return `${ARM_BASE_URL}${pathOrAbsoluteUrl}`;
}
