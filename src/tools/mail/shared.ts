import { z } from "zod";

export const Recipient = z.object({
  address: z.email(),
  name: z.string().optional(),
});

export const toRecipient = (r: z.infer<typeof Recipient>) => ({
  emailAddress: { address: r.address, name: r.name },
});
