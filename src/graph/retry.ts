import { RetryHandlerOptions, type Client as GraphClient } from "@microsoft/microsoft-graph-client";

/** Per-tool tuning of the Graph SDK's retries (429 / 503 / 504, honouring Retry-After). */
export interface RetryPolicy {
  /** Retries after the first attempt. SDK default 3, SDK maximum 10. */
  maxRetries?: number;
  /** Base delay in seconds when Graph sends no Retry-After. SDK default 3. */
  delaySeconds?: number;
}

interface RetryTarget {
  mutating?: boolean;
  retry?: RetryPolicy;
}

type ShouldRetry = ConstructorParameters<typeof RetryHandlerOptions>[2];

/**
 * Whether a failed request may be sent again.
 *
 * The SDK retries 429, 503 and 504 for every method with a JSON body. For a
 * POST that sends or creates something, only 429 is safe: Graph throttles a
 * request *before* acting on it, whereas a 503 or 504 can arrive after the mail
 * went out or the message posted — and a retry would do it twice. Reads (GET,
 * and read-only POSTs such as search or free/busy) and the repeatable writes
 * (PUT, PATCH, DELETE) keep the SDK's behaviour.
 */
export function retryDecider(tool: RetryTarget): NonNullable<ShouldRetry> {
  return (_delay, _attempt, request, options, response) => {
    const method = (typeof request === "string" ? options?.method : (request as Request).method)?.toUpperCase();
    if (tool.mutating && method === "POST") return response.status === 429;
    return true;
  };
}

export function retryOptionsFor(tool: RetryTarget): RetryHandlerOptions {
  // `undefined` lets the SDK apply its own defaults (3 s, 3 retries), which it keeps private.
  return new RetryHandlerOptions(tool.retry?.delaySeconds, tool.retry?.maxRetries, retryDecider(tool));
}

/**
 * The Graph client as one tool sees it: every request it builds carries that
 * tool's retry options, which the SDK's RetryHandler prefers over its global
 * default. Tools and the helpers they call (uploads, downloads, paging) need no
 * changes; they just use the client they are handed.
 */
export function graphForTool(graph: GraphClient, tool: RetryTarget): GraphClient {
  const options = retryOptionsFor(tool);
  return new Proxy(graph, {
    get(target, prop, receiver) {
      if (prop === "api") return (path: string) => target.api(path).middlewareOptions([options]);
      const value = Reflect.get(target, prop, receiver);
      return typeof value === "function" ? value.bind(target) : value;
    },
  });
}
