import cors from "cors";
import express, { NextFunction, Request, Response } from "express";
import { requireArmToken } from "./middleware/requireArmToken";
import { recommendationsRouter } from "./routes/recommendations";
import { scopesRouter } from "./routes/scopes";

export function createApp() {
  const app = express();

  app.use(
    cors({
      origin: process.env.FRONTEND_ORIGIN ?? true,
    }),
  );
  app.use(express.json({ limit: "1mb" }));

  app.get("/api/health", (_req, res) => {
    res.json({ status: "ok" });
  });

  app.use("/api/scopes", requireArmToken, scopesRouter);
  app.use("/api/recommendations", requireArmToken, recommendationsRouter);

  app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
    const message = error instanceof Error ? error.message : "Unknown server error";
    res.status(500).json({ error: message });
  });

  return app;
}
