export interface FilterOption {
  id: number;
  name: string;
}

export interface FilterOptionsResponse {
  meta: { user_id: number };
  data: {
    category: FilterOption[];
    semantic: FilterOption[];
    alter: FilterOption[];
  };
  request_id?: string;
}

export interface WordSearchItem {
  id: number;
  name: string;
  category_id: number;
  semantic_ids: number[];
  alter_id: number;
  lauttreu: boolean;
  image_local_url: string;
  image_external_url: string;
  image_mode: 'standard' | 'ausmalbild';
}

export interface WordSearchResponse {
  meta: {
    user_id: number;
    page: number;
    page_size: number;
    total: number;
    total_filtered: number;
  };
  data: WordSearchItem[];
  request_id?: string;
}

export interface CollectionItem {
  id: number;
  user_id: number;
  name: string;
  word_ids: number[];
}

export interface WordDetails {
  id: number;
  name: string;
  category_id: number;
  semantic_ids: number[];
  alter_id: number;
  lauttreu: boolean;
  image_local_standard_url: string;
  image_local_ausmalbild_url: string;
  image_external_url: string;
}

export interface EntitlementResponse {
  data: {
    user_id: number;
    entitled: boolean;
    plan_code: string;
    billing_period: string;
  };
  request_id?: string;
}

export interface AppConfig {
  apiBaseUrl: string;
  token: string;
}

const CONFIG_KEY = 'wortlab-word-addin-config';

function normalizeBaseUrl(value: string): string {
  return value.replace(/\/+$/, '');
}

async function request<T>(config: AppConfig, path: string, init?: RequestInit): Promise<T> {
  const headers = new Headers(init?.headers ?? {});
  headers.set('Authorization', `Bearer ${config.token}`);
  if (init?.body && !headers.has('Content-Type')) {
    headers.set('Content-Type', 'application/json');
  }

  const response = await fetch(`${normalizeBaseUrl(config.apiBaseUrl)}${path}`, {
    ...init,
    headers
  });

  if (!response.ok) {
    let message = `${response.status} ${response.statusText}`;
    try {
      const body = await response.json();
      if (body?.error) {
        message = `${message}: ${body.error}`;
      }
    } catch {
    }
    throw new Error(message);
  }

  return response.json() as Promise<T>;
}

export function loadConfig(): AppConfig {
  const raw = localStorage.getItem(CONFIG_KEY);
  if (!raw) {
    return {
      apiBaseUrl: import.meta.env.VITE_WORTLAB_API_BASE ?? '',
      token: ''
    };
  }

  try {
    const parsed = JSON.parse(raw) as Partial<AppConfig>;
    return {
      apiBaseUrl: parsed.apiBaseUrl ?? import.meta.env.VITE_WORTLAB_API_BASE ?? '',
      token: parsed.token ?? ''
    };
  } catch {
    return {
      apiBaseUrl: import.meta.env.VITE_WORTLAB_API_BASE ?? '',
      token: ''
    };
  }
}

export function saveConfig(config: AppConfig): void {
  localStorage.setItem(
    CONFIG_KEY,
    JSON.stringify({
      apiBaseUrl: normalizeBaseUrl(config.apiBaseUrl),
      token: config.token.trim()
    })
  );
}

export async function getEntitlement(config: AppConfig): Promise<EntitlementResponse> {
  return request<EntitlementResponse>(config, '/entitlement_status.php');
}

export async function getFilterOptions(config: AppConfig): Promise<FilterOptionsResponse> {
  return request<FilterOptionsResponse>(config, '/filter_options.php');
}

export async function searchWords(
  config: AppConfig,
  payload: {
    search_text: string;
    not_letter: string;
    category: number[];
    semantic: number[];
    alter: number[];
    lauttreu: boolean;
    image_mode: 'standard' | 'ausmalbild';
    page: number;
    page_size: number;
  }
): Promise<WordSearchResponse> {
  return request<WordSearchResponse>(config, '/search_words.php', {
    method: 'POST',
    body: JSON.stringify(payload)
  });
}

export async function listCollections(config: AppConfig): Promise<CollectionItem[]> {
  const result = await request<{ collections: CollectionItem[] }>(config, '/collections.php?action=list');
  return result.collections;
}

export async function getCollection(config: AppConfig, id: number): Promise<CollectionItem> {
  const result = await request<{ collection: CollectionItem }>(config, `/collections.php?action=get&id=${id}`);
  return result.collection;
}

export async function createCollection(config: AppConfig, name: string, wordIds: number[]): Promise<CollectionItem> {
  const result = await request<{ collection: CollectionItem }>(config, '/collections.php', {
    method: 'POST',
    body: JSON.stringify({ action: 'create', name, word_ids: wordIds.join(',') })
  });
  return result.collection;
}

export async function updateCollection(config: AppConfig, id: number, name: string, wordIds: number[]): Promise<CollectionItem> {
  const result = await request<{ collection: CollectionItem }>(config, '/collections.php', {
    method: 'POST',
    body: JSON.stringify({ action: 'update', id, name, word_ids: wordIds.join(',') })
  });
  return result.collection;
}

export async function getWordDetails(config: AppConfig, id: number): Promise<WordDetails> {
  const result = await request<{ data: WordDetails }>(config, `/word_details.php?id=${id}`);
  return result.data;
}
