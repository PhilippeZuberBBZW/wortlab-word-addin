import './styles.css';
import {
  createCollection,
  fetchWordImageBlob,
  getCollection,
  getEntitlement,
  getFilterOptions,
  loginWithCredentials,
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
  searchText: string;
  notLetter: string;
  lauttreuOnly: boolean;
  imageMode: 'standard' | 'ausmalbild';
  results: WordSearchItem[];
  collections: CollectionItem[];
  selectedIds: Set<number>;
  activeCollectionId: number | null;
  statusText: string;
  statusKind: 'idle' | 'error' | 'success';
  totalFiltered: number;
  entitledLabel: string;
  isEntitled: boolean;
  semanticQuery: string;
  accordionOpen: {
    category: boolean;
    alter: boolean;
    semantic: boolean;
  };
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
  searchText: '',
  notLetter: '',
  lauttreuOnly: false,
  imageMode: 'standard',
  results: [],
  collections: [],
  selectedIds: new Set<number>(),
  activeCollectionId: null,
  statusText: 'Bereit.',
  statusKind: 'idle',
  totalFiltered: 0,
  entitledLabel: 'Noch nicht geprüft',
  isEntitled: false,
  semanticQuery: '',
  accordionOpen: {
    category: false,
    alter: false,
    semantic: false
  }
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

function syncSearchStateFromForm(): void {
  state.searchText = document.querySelector<HTMLInputElement>('#searchText')?.value ?? state.searchText;
  state.notLetter = document.querySelector<HTMLInputElement>('#notLetter')?.value ?? state.notLetter;
  state.lauttreuOnly = document.querySelector<HTMLInputElement>('#lauttreu')?.checked ?? state.lauttreuOnly;
  state.imageMode = selectedImageMode();
}

function configFromForm(): AppConfig {
  const apiBaseUrl = (document.querySelector<HTMLInputElement>('#apiBaseUrl')?.value ?? '').trim();
  const token = (document.querySelector<HTMLTextAreaElement>('#accessToken')?.value ?? '').trim();
  return { apiBaseUrl, token };
}

function clearPasswordField(): void {
  const passwordInput = document.querySelector<HTMLInputElement>('#authPassword');
  if (passwordInput) {
    passwordInput.value = '';
  }
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

function filterOptionsByQuery(options: FilterOption[], query: string): FilterOption[] {
  const normalized = query.trim().toLocaleLowerCase('de-CH');
  if (!normalized) {
    return options;
  }
  return options.filter((option) => option.name.toLocaleLowerCase('de-CH').includes(normalized));
}

function renderResults(): string {
  if (state.results.length === 0) {
    return '<div class="empty">Noch keine Treffer. Führe zuerst eine Suche aus oder lade eine Sammlung.</div>';
  }

  return state.results
    .map((item) => {
      const imageUrl = getImageUrl(item);
      const checked = state.selectedIds.has(item.id) ? 'checked' : '';
      const image = imageUrl
        ? `<img class="result-preview" src="${escapeHtml(imageUrl)}" alt="${escapeHtml(item.name)}" draggable="false">`
        : '<div class="empty">Kein Bild verfügbar.</div>';

      return `
        <article class="result-card">
          <div class="result-head">
            <input class="result-checkbox" type="checkbox" data-role="select-word" data-id="${item.id}" ${checked}>
            <div>
              <h3 class="result-title">${escapeHtml(item.name)}</h3>
            </div>
          </div>
          ${image}
          <div class="result-actions">
            <button type="button" data-role="insert-text" data-id="${item.id}">Wort einfügen</button>
            <button type="button" class="secondary" data-role="insert-image" data-id="${item.id}" ${imageUrl ? '' : 'disabled'}>Bild einfügen</button>
          </div>
        </article>
      `;
    })
    .join('');
}

function renderCollections(): string {
  const options = ['<option value="">Bitte wählen...</option>']
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
  const filteredSemanticOptions = filterOptionsByQuery(state.semanticOptions, state.semanticQuery);

  app.innerHTML = `
    <main class="shell">
      <section class="hero">
        <h1>Wortlab für Word</h1>
        <p>Suche Wörter und Bilder direkt aus Wortlab und füge sie in dein Dokument ein.</p>
      </section>

      <section class="panel">
        <h2>Verbindung</h2>
        <div class="grid">
          <div class="field">
            <label for="apiBaseUrl">API-Basis</label>
            <input id="apiBaseUrl" type="url" value="${escapeHtml(state.config.apiBaseUrl)}" placeholder="https://wortlab.ch/api/v1">
          </div>
          <div class="grid two">
            <div class="field">
              <label for="authIdentifier">Benutzername oder E-Mail</label>
              <input id="authIdentifier" type="text" placeholder="z. B. name@schule.ch">
            </div>
            <div class="field">
              <label for="authPassword">Passwort</label>
              <input id="authPassword" type="password" placeholder="Passwort">
            </div>
          </div>
          <div class="field">
            <label for="accessToken">Bearer-Token (optional)</label>
            <textarea id="accessToken" placeholder="Wird nach Login automatisch gesetzt oder manuell eingefügt">${escapeHtml(state.config.token)}</textarea>
          </div>
          <div class="actions">
            <button type="button" data-role="login">Einloggen</button>
            <button type="button" data-role="save-config">Speichern</button>
            <button type="button" class="secondary" data-role="connect">Verbindung testen</button>
            <button type="button" class="ghost" data-role="logout">Ausloggen</button>
          </div>
          <div class="meta-strip">
            <span class="meta-pill">Status: ${state.config.token ? 'angemeldet' : 'abgemeldet'}</span>
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
            <input id="searchText" type="text" value="${escapeHtml(state.searchText)}" placeholder="z. B. *le oder ba*">
          </div>
          <div class="grid two">
            <div class="field">
              <label for="notLetter">Buchstabe ausschliessen</label>
              <input id="notLetter" type="text" value="${escapeHtml(state.notLetter)}" maxlength="10" placeholder="z. B. r">
            </div>
            <div class="field">
              <label for="imageMode">Bildmodus</label>
              <select id="imageMode">
                <option value="standard" ${state.imageMode === 'standard' ? 'selected' : ''}>Standard</option>
                <option value="ausmalbild" ${state.imageMode === 'ausmalbild' ? 'selected' : ''}>Ausmalbild</option>
              </select>
            </div>
          </div>
          <label class="checkline"><input id="lauttreu" type="checkbox" ${state.lauttreuOnly ? 'checked' : ''}> <span>Lauttreu</span></label>
          <div class="accordion-group">
            <details class="accordion" data-accordion="category" ${state.accordionOpen.category ? 'open' : ''}>
              <summary>Wortarten</summary>
              <div class="accordion-content grid">${renderOptions('category', state.categoryOptions, selectedCategory)}</div>
            </details>
            <details class="accordion" data-accordion="alter" ${state.accordionOpen.alter ? 'open' : ''}>
              <summary>Alter</summary>
              <div class="accordion-content grid">${renderOptions('alter', state.alterOptions, selectedAlter)}</div>
            </details>
            <details class="accordion" data-accordion="semantic" ${state.accordionOpen.semantic ? 'open' : ''}>
              <summary>Kategorien</summary>
              <div class="accordion-content grid">
                <div class="field">
                  <label for="semanticFilter">Kategorien filtern</label>
                  <input id="semanticFilter" type="search" value="${escapeHtml(state.semanticQuery)}" placeholder="Kategorie suchen ...">
                </div>
                ${renderOptions('semantic', filteredSemanticOptions, selectedSemantic)}
              </div>
            </details>
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

function renderPreservingView(focusElementId?: string): void {
  const scrollX = window.scrollX;
  const scrollY = window.scrollY;
  const activeElement = document.activeElement;

  let activeId: string | null = null;
  let selectionStart: number | null = null;
  let selectionEnd: number | null = null;

  if (activeElement instanceof HTMLInputElement || activeElement instanceof HTMLTextAreaElement) {
    activeId = activeElement.id;
    selectionStart = activeElement.selectionStart;
    selectionEnd = activeElement.selectionEnd;
  }

  render();
  window.scrollTo(scrollX, scrollY);

  const nextFocusId = focusElementId ?? activeId;
  if (!nextFocusId) {
    return;
  }

  const nextFocusElement = document.getElementById(nextFocusId);
  if (!(nextFocusElement instanceof HTMLInputElement || nextFocusElement instanceof HTMLTextAreaElement)) {
    return;
  }

  nextFocusElement.focus({ preventScroll: true });
  if (selectionStart !== null && selectionEnd !== null) {
    nextFocusElement.setSelectionRange(selectionStart, selectionEnd);
  }
}

async function connect(): Promise<void> {
  state.config = configFromForm();
  saveConfig(state.config);

  if (!state.config.apiBaseUrl) {
    setStatus('Bitte zuerst die API-Basis eintragen.', 'error');
    return;
  }
  if (!state.config.token) {
    setStatus('Bitte zuerst einloggen oder einen Bearer-Token eintragen.', 'error');
    return;
  }

  setStatus('Verbindung wird geprüft ...');

  const [entitlement, filters, collections] = await Promise.all([
    getEntitlement(state.config),
    getFilterOptions(state.config),
    listCollections(state.config)
  ]);

  state.entitledLabel = entitlement.data.entitled
    ? `${entitlement.data.plan_code} · ${entitlement.data.billing_period}`
    : 'kein Zugang';
  state.isEntitled = entitlement.data.entitled;
  state.categoryOptions = filters.data.category;
  state.semanticOptions = filters.data.semantic;
  state.alterOptions = filters.data.alter;
  state.collections = collections;
  if (!state.isEntitled) {
    setStatus('Login erfolgreich, aber kein aktives Abo. Bitte Freischaltung ausserhalb des Add-ins veranlassen.', 'error');
    return;
  }
  setStatus('Verbindung erfolgreich. Filter und Sammlungen geladen.', 'success');
}

async function loginDirect(): Promise<void> {
  const apiBaseUrl = (document.querySelector<HTMLInputElement>('#apiBaseUrl')?.value ?? '').trim();
  const identifier = (document.querySelector<HTMLInputElement>('#authIdentifier')?.value ?? '').trim();
  const password = document.querySelector<HTMLInputElement>('#authPassword')?.value ?? '';

  if (!apiBaseUrl) {
    setStatus('Bitte zuerst die API-Basis eintragen.', 'error');
    return;
  }
  if (!identifier || !password) {
    setStatus('Bitte Benutzername/E-Mail und Passwort eingeben.', 'error');
    return;
  }

  setStatus('Login läuft ...');
  const login = await loginWithCredentials(apiBaseUrl, identifier, password);

  state.config = {
    apiBaseUrl,
    token: login.token
  };
  saveConfig(state.config);
  clearPasswordField();
  render();

  await connect();
}

function logoutDirect(): void {
  state.config = {
    apiBaseUrl: (document.querySelector<HTMLInputElement>('#apiBaseUrl')?.value ?? state.config.apiBaseUrl).trim(),
    token: ''
  };
  saveConfig(state.config);
  state.entitledLabel = 'Noch nicht geprüft';
  state.isEntitled = false;
  state.results = [];
  state.totalFiltered = 0;
  state.selectedIds = new Set<number>();
  clearPasswordField();
  render();
  setStatus('Du bist ausgeloggt.', 'success');
}

async function runSearch(): Promise<void> {
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Suche ist gesperrt.', 'error');
    return;
  }

  state.config = configFromForm();
  saveConfig(state.config);
  syncSearchStateFromForm();
  setStatus('Suche läuft ...');

  const response = await searchWords(state.config, {
    search_text: state.searchText,
    not_letter: state.notLetter,
    category: selectedValues('category'),
    semantic: selectedValues('semantic'),
    alter: selectedValues('alter'),
    lauttreu: state.lauttreuOnly,
    image_mode: state.imageMode,
    page: 1,
    page_size: 25
  });

  state.results = response.data;
  state.totalFiltered = response.meta.total_filtered;
  setStatus(`${response.meta.total_filtered} Treffer geladen.`, 'success');
}

async function loadSelectedCollection(): Promise<void> {
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Sammlungen sind gesperrt.', 'error');
    return;
  }

  const id = Number(document.querySelector<HTMLSelectElement>('#collectionSelect')?.value ?? '0');
  if (!id) {
    setStatus('Bitte zuerst eine Sammlung wählen.', 'error');
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
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Sammlungen sind gesperrt.', 'error');
    return;
  }

  const id = Number(document.querySelector<HTMLSelectElement>('#collectionSelect')?.value ?? '0');
  if (!id) {
    setStatus('Bitte zuerst eine bestehende Sammlung wählen.', 'error');
    return;
  }

  const name = document.querySelector<HTMLSelectElement>('#collectionSelect')?.selectedOptions[0]?.textContent?.trim() ?? '';
  if (!name) {
    setStatus('Sammlung konnte nicht gelesen werden.', 'error');
    return;
  }

  setStatus('Sammlung wird gespeichert ...');
  const updated = await updateCollection(state.config, id, name, Array.from(state.selectedIds));
  state.collections = state.collections.map((c) => (c.id === id ? updated : c));
  state.activeCollectionId = id;
  setStatus('Sammlung aktualisiert.', 'success');
}

async function createNewCollection(): Promise<void> {
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Sammlungen sind gesperrt.', 'error');
    return;
  }

  const name = (document.querySelector<HTMLInputElement>('#collectionName')?.value ?? '').trim();
  if (!name) {
    setStatus('Bitte einen Namen für die neue Sammlung eingeben.', 'error');
    return;
  }

  setStatus('Sammlung wird erstellt ...');
  const collection = await createCollection(state.config, name, Array.from(state.selectedIds));
  state.collections = [...state.collections, collection];
  state.activeCollectionId = collection.id;
  setStatus(`Sammlung \"${collection.name}\" erstellt.`, 'success');
}

function findWordById(id: number): WordSearchItem | undefined {
  return state.results.find((item) => item.id === id);
}

async function handleInsertText(id: number): Promise<void> {
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Einfügen ist gesperrt.', 'error');
    return;
  }

  const item = findWordById(id);
  if (!item) {
    setStatus('Wort nicht gefunden.', 'error');
    return;
  }

  setStatus(`\"${item.name}\" wird in Word eingefügt ...`);
  await insertWordText(item.name);
  setStatus(`\"${item.name}\" wurde eingefügt.`, 'success');
}

async function handleInsertImage(id: number): Promise<void> {
  if (!state.isEntitled) {
    setStatus('Keine aktive Berechtigung. Einfügen ist gesperrt.', 'error');
    return;
  }

  const item = findWordById(id);
  const imageUrl = item ? getImageUrl(item) : '';
  if (!item || !imageUrl) {
    setStatus('Kein Bild für dieses Wort verfügbar.', 'error');
    return;
  }

  setStatus(`Bild zu \"${item.name}\" wird eingefügt ...`);
  const imageBlob = await fetchWordImageBlob(state.config, item.id, state.imageMode);
  await insertWordImage(imageBlob);
  setStatus(`Bild zu \"${item.name}\" wurde eingefügt.`, 'success');
}

function toggleSelection(id: number, checked: boolean): void {
  if (checked) {
    state.selectedIds.add(id);
  } else {
    state.selectedIds.delete(id);
  }
  render();
}

async function handleAction(actionElement: HTMLElement): Promise<void> {
  const role = actionElement.dataset.role;
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

    if (role === 'login') {
      await loginDirect();
      return;
    }

    if (role === 'logout') {
      logoutDirect();
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

    const id = Number(actionElement.dataset.id ?? '0');
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
  const target = event.target;
  if (!(target instanceof Element)) {
    return;
  }
  const actionElement = target.closest<HTMLElement>('[data-role]');
  if (!actionElement) {
    return;
  }
  void handleAction(actionElement);
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

app.addEventListener('input', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }

  if (target.id === 'semanticFilter') {
    state.semanticQuery = target.value;
    renderPreservingView('semanticFilter');
  }
});

app.addEventListener('toggle', (event) => {
  const target = event.target;
  if (!(target instanceof HTMLDetailsElement)) {
    return;
  }

  const accordionKey = target.dataset.accordion;
  if (accordionKey === 'category' || accordionKey === 'alter' || accordionKey === 'semantic') {
    state.accordionOpen[accordionKey] = target.open;
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
