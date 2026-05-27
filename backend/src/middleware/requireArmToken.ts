import { NextFunction, Request, Response } from "express";

export function requireArmToken(req: Request, res: Response, next: NextFunction): void {
  const authHeader = req.header("Authorization");
  if (!authHeader || !authHeader.toLowerCase().startsWith("bearer ")) {
    res.status(401).json({
      error:
        "Missing Bearer token. Provide an Azure Management API access token in Authorization header.",
    });
    return;
  }

  next();
}

export function getBearerToken(req: Request): string {
  const authHeader = req.header("Authorization") ?? "";
  return authHeader.replace(/^Bearer\s+/i, "").trim();
}
