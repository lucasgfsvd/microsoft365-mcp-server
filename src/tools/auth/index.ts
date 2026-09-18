import { z } from "zod";
import type { ToolDefinition } from "../../types.js";

const NoInput = z.object({});

export const authTools: ToolDefinition[] = [
  {
    name: "auth_status",
    surface: "auth",
    description:
      "Report whether the server is signed in to Microsoft 365, and surface any device code " +
      "that is currently waiting to be entered.",
    inputSchema: NoInput,
    handler: async (_input, ctx) => ctx.auth.status(),
  },
  {
    name: "auth_sign_in",
    surface: "auth",
    // Deliberately NOT marked `mutating`: sign-in has to stay reachable even when
    // writes are disabled, otherwise a read-only deployment could never authenticate.
    description:
      "Start an interactive Microsoft 365 sign-in and return the device code to enter. " +
      "Returns immediately with the code; call auth_status afterwards to confirm it completed.",
    inputSchema: NoInput,
    handler: async (_input, ctx) => ctx.auth.signIn(),
  },
];
