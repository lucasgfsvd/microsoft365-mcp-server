import { z } from "zod";
import type { ToolDefinition } from "../../types.js";
import { fetchPage } from "../../graph/pagination.js";
import { PaginationInput } from "../../util/schema.js";

/** The organisation's bookable rooms; Place.Read.All needs an admin's approval. */
export const roomTools: ToolDefinition[] = [
  {
    name: "calendar_list_rooms",
    surface: "calendar",
    description:
      "List the organisation's meeting rooms, with capacity, building and floor. Check which are free with " +
      "calendar_get_free_busy on their addresses, then book one with calendar_create_event's room.",
    requiredScopes: ["Place.Read.All"],
    needsAdminConsent: true,
    inputSchema: PaginationInput.extend({
      building: z.string().optional().describe("Only rooms in this building."),
      minCapacity: z.number().int().min(1).optional(),
    }),
    handler: async (input, ctx) => {
      const page = await fetchPage<Record<string, unknown>>(ctx.graph, `/places/microsoft.graph.room`, input);
      const value = page.value
        .filter((r) => !input.building || String(r.building ?? "").toLowerCase() === input.building.toLowerCase())
        .filter((r) => !input.minCapacity || Number(r.capacity ?? 0) >= input.minCapacity)
        .map((r) => ({
          name: r.displayName,
          email: r.emailAddress,
          capacity: r.capacity,
          building: r.building,
          floor: r.floorNumber ?? r.floorLabel,
          features: [r.isWheelChairAccessible && "wheelchair access", r.videoDeviceName && "video", r.displayDeviceName && "display"].filter(Boolean),
        }));
      return { ...page, value };
    },
  },
];
