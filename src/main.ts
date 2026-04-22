import './styles.css';
import {
  createCollection,
  getCollection,
  getEntitlement,
  getFilterOptions,
  getWordDetails,
  listCollections,
  loadConfig,
  saveConfig,
  searchWords,
  updateCollection,
  type AppConfig,
  type CollectionItem,
  type FilterOption,
  type WordSearchItem
} from './api';
import { insertWordImage, insertWordText } from './office';

interface AppState {
  config: AppConfig;
  categoryOptions: FilterOption[];
  semanticOptions: FilterOption[];
  alterOptions: FilterOption[];
  results: WordSearchItem[];
  collections: CollectionItem[];
  selectedIds: Set<number>;
  activeCollectionId: number | null;
  statusText: string;
  statusKind: 'idle' | 'error' | 'success';
  totalFiltered: number;
  entitledLabel: string;
}

const appElement = document.querySelector<HTMLDivElement>('#app');
if (!appElement) {
  throw new Error('App root not found.');
}
const app: HTMLDivElement = appElement;

const state: AppState = {
  config: loadConfig(),
  categoryOptions: [],
  semanticOptions: [],
  alterOptions: [],
  results: [],
  collections: [],
  selectedIds: new Set<number>(),
  activeCollectionId: null,
  statusText: 'Bereit.',
  statusKind: 'idle',
  totalFiltered: 0,
  entitledLabel: 'Noch nicht geprueft'
};

function setStatus(text: string, kind: AppState['statusKind'] = 'idle'): void {
  state.statusText = text;
  state.statusKind = kind;
  render();
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getImageUrl(item: WordSearchItem): string {
  return item.image_local_url || item.image_external_url || '';
}

function selectedValues(name: string): number[] {
  return Array.from(document.querySelectorAll<HTMLInputElement>(`input[name="${name}"]:checked`)).map((input) => Number(input.value));
}

function selectedImageMode(): 'standard' | 'ausmalbild' {
  const select = document.querySelector<HTMLSelectElement>('#imageMode');
  return select?.value === 'ausmalbild' ? 'ausmalbild' : 'standard';
}

function configFromForm(): AppConfig {
  const apiBaseUrl = (document.querySelector<HTMLInputElement>('#apiBaseUrl')?.value ?? '').trim();
  const token = (document.querySelector<HTMLTextAreaElement>('#accessToken')?.value ?? '').trim();
  return { apiBaseUrl, token };
}

function renderOptions(name: string, options: FilterOption[], selected: Set<number>): string {
  if (options.length === 0) {
    return '<div class="note">Noch nicht geladen.</div>';
  }

  return options
    .map((option) => {
      const checked = selected.has(option.id) ? 'checked' : '';
      return `<label class="checkline"><input type="checkbox" name="${name}" value="${option.id}" ${checked}> <span>${escapeHtml(option.name)}</span></label>`;
    })
    .join('');
}

function renderResults(): string {
  if (state.results.length === 0) {
    return '<div class="empty">Noch keine Treffer. Fuehre zuerst eine Suche aus oder lade eine Sammlung.</div>';
  }

  return state.results
    .map((item) => {
      const imageUrl = getImageUrl(item);
      const checked = state.selectedIds.has(item.id) ? 'checked' : '';
      const image = imageUrl
        ? `<img class="result-preview" src="${escapeHtml(imageUrl)}" alt="${escapeHtml(item.name)}">`
        : '<div class="empty">Kein Bild verfuegbar.</div>';

      return `
        <article class="result-card">
          <div class="result-head">
            <input class="result-checkbox" type="checkbox" data-role="select-word" data-id="${item.id}" ${checked}>
            <div>
              <h3 class="result-title">${escapeHtml(item.name)}</h3>
              <p class="result-sub">ID ${item.id} · ${item.lauttreu ? 'lauttreu' : 'nicht lauttreu'}</p>
            </div>
          </div>
          ${image}
          <div class="result-actions">
            <button type="button" data-role="insert-text" data-id="${item.id}">Wort einfuegen</button>
            <button type="button" class="secondary" data-role="insert-image" data-id="${item.id}" ${imageUrl ? '' : 'disabled'}>Bild einfuegen</button>
          </div>
        </article>
      `;
    })
    .join('');
}

function renderCollections(): string {
  const options = ['<option value="">Bitte waehlen...</option>']
    .concat(
      state.collections.map((collection) => {
        const selected = state.activeCollectionId === collection.id ? 'selected' : '';
        return `<option value="${collection.id}" ${selected}>${escapeHtml(collection.name)}</option>`;
      })
    )
    .join('');

  return `
    <div class="collection-row">
      <div class="field">
        <label for="collectionSelect">Sammlung</label>
        <select id="collectionSelect">${options}</select>
      </div>
      <div class="actions">
        <button type="button" data-role="load-collection">Laden</button>
        <button type="button" class="secondary" data-role="save-collection">Auswahl speichern</button>
      </div>
      <div class="field">
        <label for="collectionName">Neue oder umbenannte Sammlung</label>
        <input id="collectionName" type="text" placeholder="z. B. S-Laute Woche 3">
      </div>
      <div class="actions">
        <button type="button" class="ghost" data-role="create-collection">Neu anlegen</button>
      </div>
    </div>
  `;
}

function render(): void {
  const selectedCategory = new Set(selectedValues('category'));
  const selectedSemantic = new Set(selectedValues('semantic'));
  const selectedAlter = new Set(selectedValues('alter'));

  app.innerHTML = `
    <main class="shell">
      <section class="hero">
        <h1>Wortlab fuer Word</h1>
        <p>Suche Woerter und Bilder direkt aus Wortlab und fuege sie in dein Dokument ein.</p>
      </section>

      <section class="panel">
        <h2>Verbindung</h2>
        <div class="grid">
          <div class="field">
            <label for="apiBaseUrl">API-Basis</label>
            <input id="apiBaseUrl" type="url" value="${escapeHtml(state.config.apiBaseUrl)}" placeholder="https://wortlab.ch/api/v1">
          </div>
          <div class="field">
            <label for="accessToken">Bearer-Token</label>
            <textarea id="accessToken" placeholder="Token aus auth_token.php">${escapeHtml(state.config.token)}</textarea>
          </div>
          <div class="actions">
            <button type="button" data-role="save-config">Speichern</button>
            <button type="button" class="secondary" data-role="connect">Verbindung testen</button>
          </div>
          <div class="meta-strip">
            <span class="meta-pill">Entitlement: ${escapeHtml(state.entitledLabel)}</span>
            <span class="meta-pill">Treffer: ${state.totalFiltered}</span>
            <span class="meta-pill">Auswahl: ${state.selectedIds.size}</span>
          </div>
          <div class="status ${state.statusKind === 'error' ? 'error' : state.statusKind === 'success' ? 'success' : ''}">${escapeHtml(state.statusText)}</div>
        </div>
      </section>

      <section class="panel">
        <h2>Suche</h2>
        <div class="grid">
          <div class="field">
            <label for="searchText">Suchtext</label>
            <input id="searchText" type="text" placeholder="z. B. *le oder ba*">
          </div>
          <div class="grid two">
            <div class="field">
              <label for="notLetter">Buchstabe ausschliessen</label>
              <input id="notLetter" type="text" maxlength="10" placeholder="z. B. r">
            </div>
            <div class="field">
              <label for="imageMode">Bildmodus</label>
              <select id="imageMode">
                <option value="standard">Standard</option>
                <option value="ausmalbild">Ausmalbild</option>
              </select>
            </div>
          </div>
          <label class="checkline"><input id="lauttreu" type="checkbox"> <span>Lauttreu</span></label>
          <div class="grid two">
            <div class="field">
              <span class="label">Wortarten</span>
              <div class="grid">${renderOptions('category', state.categoryOptions, selectedCategory)}</div>
            </div>
            <div class="field">
              <span class="label">Alter</span>
              <div class="grid">${renderOptions('alter', state.alterOptions, selectedAlter)}</div>
            </div>
          </div>
          <div class="field">
            <span class="label">Kategorien</span>
            <div class="grid">${renderOptions('semantic', state.semanticOptions, selectedSemantic)}</div>
          </div>
          <div class="actions">
            <button type="button" data-role="search">Suchen</button>
          </div>
          <div class="note">Sternchen-Suche: abc*, *abc, *abc* und abc werden an die Wortlab-API weitergegeben.</div>
        </div>
      </section>

      <section class="panel">
        <h2>Wortsammlungen</h2>
        ${renderCollections()}
      </section>

      <section class="panel">
        <h2>Trefferliste</h2>
        <div class="result-list">${renderResults()}</div>
      </section>
    </main>
  `;
}

async function connect(): Promise<void> {
  state.config = configFromForm();
  saveConfig(state.config);
  setStatus('Verbindung wird geprueft ...');

  const [entitlement, filters, collections] = await Promise.all([
    getEntitlement(state.config),
    getFilterOptions(state.config),
    listCollections(state.config)
  ]);

  state.entitledLabel = entitlement.data.entitled
    ? `${entitlement.data.plan_code} · ${entitlement.data.billing_period}`
    : 'kein Zugang';
  state.categoryOptions = filters.data.category;
  state.semanticOptions = filters.data.semantic;
  state.alterOptions = filters.data.alter;
  state.collections = collections;
  setStatus('Verbindung erfolgreich. Filter und Sammlungen geladen.', 'success');
}

async function runSearch(): Promise<void> {
  state.config = configFromForm();
  saveConfig(state.config);
  setStatus('Suche laeuft ...');

  const response = await searchWords(state.config, {
    search_text: document.querySelector<HTMLInputElement>('#searchText')?.value ?? '',
    not_letter: document.querySelector<HTMLInputElement>('#notLetter')?.value ?? '',
    category: selectedValues('category'),
    semantic: selectedValues('semantic'),
    alter: selectedValues('alter'),
    lauttreu: document.querySelector<HTMLInputElement>('#lauttreu')?.checked ?? false,
    image_mode: selectedImageMode(),
    page: 1,
    page_size: 25
  });

  state.results = response.data;
  state.totalFiltered = response.meta.total_filtered;
  setStatus(`${response.meta.total_filtered} Treffer geladen.`, 'success');
}

async function loadSelectedCollection(): Promise<void> {
  const id = Number(document.querySelector<HTMLSelectElement>('#collectionSelect')?.value ?? '0');
  if (!id) {
    setStatus('Bitte zuerst eine Sammlung waehlen.', 'error');
    return;
  }

  setStatus('Sammlung wird geladen ...');
  const collection = await getCollection(state.config, id);
  state.activeCollectionId = collection.id;
  state.selectedIds = new Set(collection.word_ids);
  const details = await Promise.all(collection.word_ids.map((wordId) => getWordDetails(state.config, wordId)));
  state.results = details.map((item) => ({
    id: item.id,
    name: item.name,
    category_id: item.category_id,
    semantic_ids: item.semantic_ids,
    alter_id: item.alter_id,
    lauttreu: item.lauttreu,
    image_local_url: item.image_local_standard_url,
    image_external_url: item.image_external_url,
    image_mode: selectedImageMode()
  }));
  state.totalFiltered = state.results.length;
  setStatus(`Sammlung \"${collection.name}\" geladen.`, 'success');
}

async function saveCurrentSelectionToCollection(): Promise<void> {
  const id = Number(document.querySelector<HTMLSelectElement>('#collectionSelect')?.value ?? '0');
  if (!id) {
    setStatus('Bitte zuerst eine bestehende Sammlung waehlen.', 'error');
    return;
  }

  const name = document.querySelector<HTMLSelectElement>('#collectionSelect')?.selectedOptions[0]?.textContent?.trim() ?? '';
  if (!name) {
    setStatus('Sammlung konnte nicht gelesen werden.', 'error');
    return;
  }

  setStatus('Sammlung wird gespeichert ...');
  await updateCollection(state.config, id, name, Array.from(state.selectedIds));
  state.collections = await listCollections(state.config);
  state.activeCollectionId = id;
  setStatus('Sammlung aktualisiert.', 'success');
}

async function createNewCollection(): Promise<void> {
  const name = (document.querySelector<HTMLInputElement>('#collectionName')?.value ?? '').trim();
  if (!name) {
    setStatus('Bitte einen Namen fuer die neue Sammlung eingeben.', 'error');
    return;
  }

  setStatus('Sammlung wird erstellt ...');
  const collection = await createCollection(state.config, name, Array.from(state.selectedIds));
  state.collections = await listCollections(state.config);
  state.activeCollectionId = collection.id;
  setStatus(`Sammlung \"${collection.name}\" erstellt.`, 'success');
}

function findWordById(id: number): WordSearchItem | undefined {
  return state.results.find((item) => item.id === id);
}

async function handleInsertText(id: number): Promise<void> {
  const item = findWordById(id);
  if (!item) {
    setStatus('Wort nicht gefunden.', 'error');
    return;
  }

  setStatus(`\"${item.name}\" wird in Word eingefuegt ...`);
  await insertWordText(item.name);
  setStatus(`\"${item.name}\" wurde eingefuegt.`, 'success');
}

async function handleInsertImage(id: number): Promise<void> {
  const item = findWordById(id);
  const imageUrl = item ? getImageUrl(item) : '';
  if (!item || !imageUrl) {
    setStatus('Kein Bild fuer dieses Wort verfuegbar.', 'error');
    return;
  }

  setStatus(`Bild zu \"${item.name}\" wird eingefuegt ...`);
  await insertWordImage(imageUrl);
  setStatus(`Bild zu \"${item.name}\" wurde eingefuegt.`, 'success');
}

function toggleSelection(id: number, checked: boolean): void {
  if (checked) {
    state.selectedIds.add(id);
  } else {
    state.selectedIds.delete(id);
  }
  render();
}

async function handleAction(target: HTMLElement): Promise<void> {
  const role = target.dataset.role;
  if (!role) {
    return;
  }

  try {
    if (role === 'save-config') {
      state.config = configFromForm();
      saveConfig(state.config);
      setStatus('Konfiguration gespeichert.', 'success');
      return;
    }

    if (role === 'connect') {
      await connect();
      return;
    }

    if (role === 'search') {
      await runSearch();
      return;
    }

    if (role === 'load-collection') {
      await loadSelectedCollection();
      return;
    }

    if (role === 'save-collection') {
      await saveCurrentSelectionToCollection();
      return;
    }

    if (role === 'create-collection') {
      await createNewCollection();
      return;
    }

    const id = Number(target.dataset.id ?? '0');
    if (!id) {
      return;
    }

    if (role === 'insert-text') {
      await handleInsertText(id);
      return;
    }

    if (role === 'insert-image') {
      await handleInsertImage(id);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Unbekannter Fehler';
    setStatus(message, 'error');
  }
}

app.addEventListener('click', (event) => {
  const target = event.target as HTMLElement | null;
  if (!target) {
    return;
  }
  void handleAction(target);
});

app.addEventListener('change', (event) => {
  const target = event.target as HTMLElement | null;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  if (target.dataset.role === 'select-word') {
    const id = Number(target.dataset.id ?? '0');
    if (id) {
      toggleSelection(id, target.checked);
    }
  }
});

async function bootstrap(): Promise<void> {
  render();
  await Office.onReady();
  if (state.config.apiBaseUrl && state.config.token) {
    try {
      await connect();
    } catch (error) {
      const message = error instanceof Error ? error.message : 'Verbindung fehlgeschlagen';
      setStatus(message, 'error');
    }
  } else {
    setStatus('API-Basis und Token eintragen, dann Verbindung testen.');
  }
}

void bootstrap();
