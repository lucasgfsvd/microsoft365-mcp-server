import type { Client as GraphClient } from "@microsoft/microsoft-graph-client";

export interface RecentFile {
  id: string;
  driveId?: string;
  name?: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: string;
  webUrl?: string;
}

interface Hit {
  resource?: {
    id?: string;
    name?: string;
    webUrl?: string;
    lastModifiedDateTime?: string;
    lastModifiedBy?: { user?: { displayName?: string } };
    parentReference?: { driveId?: string };
  };
}

/**
 * The documents the user can reach that changed most recently, newest first.
 *
 * Microsoft retires /me/drive/recent (and sharedWithMe, and the file insights)
 * after November 2026, so this asks Microsoft Search instead: documents,
 * ordered by last change. It covers SharePoint as well as OneDrive.
 */
export async function recentFiles(graph: GraphClient, size = 20): Promise<RecentFile[]> {
  const res = (await graph.api("/search/query").post({
    requests: [
      {
        entityTypes: ["driveItem"],
        query: { queryString: "IsDocument:true" },
        sortProperties: [{ name: "lastModifiedDateTime", isDescending: true }],
        from: 0,
        size,
      },
    ],
  })) as { value?: Array<{ hitsContainers?: Array<{ hits?: Hit[] }> }> };
  const hits = res.value?.[0]?.hitsContainers?.[0]?.hits ?? [];
  return hits.flatMap(({ resource: r }) =>
    r?.id
      ? [{
          id: r.id,
          driveId: r.parentReference?.driveId,
          name: r.name,
          lastModifiedDateTime: r.lastModifiedDateTime,
          lastModifiedBy: r.lastModifiedBy?.user?.displayName,
          webUrl: r.webUrl,
        }]
      : [],
  );
}
