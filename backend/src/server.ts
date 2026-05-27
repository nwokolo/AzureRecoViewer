import dotenv from "dotenv";
import { createApp } from "./app";

dotenv.config();

const app = createApp();
const port = Number.parseInt(process.env.PORT ?? "4000", 10);

app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`Backend listening on port ${port}`);
});
