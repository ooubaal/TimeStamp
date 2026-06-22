export interface GistSyncConfig {
  githubToken: string;
  gistId: string;
}

// Key used in localStorage
const GIST_CONFIG_KEY = "TimeStampGistConfig";

export function loadGistConfig(): GistSyncConfig | null {
  const stored = localStorage.getItem(GIST_CONFIG_KEY);
  if (stored) {
    try {
      return JSON.parse(stored);
    } catch (e) {
      console.error("Failed to parse Gist config", e);
    }
  }
  return null;
}

export function saveGistConfig(config: GistSyncConfig) {
  localStorage.setItem(GIST_CONFIG_KEY, JSON.stringify(config));
}

export function clearGistConfig() {
  localStorage.removeItem(GIST_CONFIG_KEY);
}

/**
 * Fetch settings JSON from GitHub Gist
 */
export async function fetchFromGist(config: GistSyncConfig): Promise<any> {
  const response = await fetch(`https://api.github.com/gists/${config.gistId}`, {
    headers: {
      Authorization: `token ${config.githubToken}`,
      Accept: "application/vnd.github.v3+json",
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to fetch Gist: ${response.statusText} (${response.status})`);
  }

  const gistData = await response.json();
  const fileKey = Object.keys(gistData.files).find(key => key.endsWith(".json"));
  
  if (!fileKey) {
    throw new Error("No JSON database file found in the specified Gist");
  }

  const fileContent = gistData.files[fileKey].content;
  return JSON.parse(fileContent);
}

/**
 * Update Gist with new settings JSON database
 */
export async function updateGist(config: GistSyncConfig, dbData: any): Promise<void> {
  const body = {
    description: "TimeStamp Analysis DB Configuration (Sync)",
    files: {
      "TimeStampDB.json": {
        content: JSON.stringify(dbData, null, 2),
      },
    },
  };

  const response = await fetch(`https://api.github.com/gists/${config.gistId}`, {
    method: "PATCH",
    headers: {
      Authorization: `token ${config.githubToken}`,
      Accept: "application/vnd.github.v3+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    throw new Error(`Failed to update Gist: ${response.statusText} (${response.status})`);
  }
}
