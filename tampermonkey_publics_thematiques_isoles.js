/**
 * MODULE ISOLÉ — AJOUT DES PUBLICS ET DES THÉMATIQUES
 * DANS UNE FICHE FORMATION DRUPAL
 *
 * Ce module :
 * 1. ajoute plusieurs thématiques dans le champ d’autocomplétion Drupal ;
 * 2. sélectionne plusieurs publics/métiers dans l’arbre Fancytree Drupal.
 *
 * Important :
 * Il ne crée pas de nouveaux termes de taxonomie Drupal.
 * Les thématiques et les publics doivent déjà exister dans Drupal.
 *
 * Exemple d’utilisation :
 *
 * await window.eafcPublicThematicTools.applyAll({
 *   item: {
 *     theme: 'Numérique, École inclusive',
 *     publics: [
 *       'Tous les personnels',
 *       {
 *         path: 'Tous les personnels > Personnels enseignants > Tous les enseignants'
 *       }
 *     ]
 *   },
 *   exportData: {
 *     theme: '',
 *     publics: []
 *   }
 * });
 */

(function () {
  'use strict';

  // ---------------------------------------------------------------------------
  // 1. CONFIGURATION
  // ---------------------------------------------------------------------------

  const CONFIG = {
    delayMs: 900,
    safeMode: false,
    autocompleteDelayMs: 800,
    autocompleteSelectionDelayMs: 600,

    // Thématiques systématiquement ajoutées à chaque fiche.
    requiredThematics: ['EAFC poitiers', '6331'],

    // Correspondance entre les libellés reçus et les identifiants Drupal.
    thematicMap: [
      [
        [
          'Pédagogique',
          'Pratique pédagogique transversales',
          'Pratiques pédagogiques transversales',
        ],
        ['6624'],
      ],
      [
        [
          'Numérique',
          'Compétences, culture et usages du numérique',
        ],
        ['6362'],
      ],
      [
        [
          'QVCT, SST',
          'QVCT',
          'SST',
          'Qualité de vie au travail – Santé et sécurité au travail',
        ],
        ['6385', '6404'],
      ],
      [
        [
          'Valeurs de la république, citoyenneté',
          'Valeurs de la République et citoyenneté',
          'Valeurs de la république',
          'citoyenneté',
        ],
        ['6627', '6236'],
      ],
      [
        [
          'Orientation',
          'Orientation et parcours scolaire',
        ],
        ['6364'],
      ],
      [
        [
          'Ecole inclusive',
          'École inclusive',
          'Ecole inclusive, accessible et ouverte à tous',
        ],
        ['6314'],
      ],
      [
        [
          'Europe et international',
          'Ouverture européenne et internationale',
        ],
        ['6328'],
      ],
      [
        [
          'Formateurs',
          'Formation de formateurs et tuteurs',
        ],
        ['6625'],
      ],
      [
        [
          'Santé des élèves',
          'Santé mentale et bien-être des élèves',
        ],
        ['6403'],
      ],
      [
        ['Encadrement'],
        ['6455'],
      ],
      [
        [
          'Carrière',
          'Carrière et évolution professionnelle',
        ],
        ['6227'],
      ],
      [
        [
          '2nd degré',
          'Second degré',
        ],
        ['6202'],
      ],
    ],
  };

  const SELECTORS = {
    // Champ Entity autocomplete des thématiques.
    thematicInput:
      'input[name^="field_formation_thematic"][name$="[target_id]"], ' +
      'input[data-drupal-selector^="edit-field-formation-thematic-"]' +
      '[data-autocomplete-path]',

    // Bouton Drupal « Ajouter un autre élément ».
    thematicAddMore:
      '[data-drupal-selector="edit-field-formation-thematic-add-more"], ' +
      'input[name="field_formation_thematic_add_more"]',

    // Conteneur Fancytree du champ publics/métiers.
    publicTreeWrapper: '#edit-field-metier-tags-wrapper',
  };

  // ---------------------------------------------------------------------------
  // 2. API PRINCIPALE
  // ---------------------------------------------------------------------------

  /**
   * Ajoute les thématiques puis les publics.
   */
  async function applyAll({
    item = {},
    exportData = {},
  } = {}) {
    const thematics = await applyThematics({
      item,
      exportData,
    });

    const publics = await applyPublics({
      item,
      exportData,
    });

    return {
      thematics,
      publics,
    };
  }

  /**
   * Ajoute toutes les thématiques résolues
   * dans les champs Drupal.
   */
  async function applyThematics({
    item = {},
    exportData = {},
  } = {}) {
    const thematics = resolveThematics({
      item,
      exportData,
    });

    if (!thematics.length) {
      log('Aucune thématique à ajouter.');
      return [];
    }

    log(`Thématiques à ajouter : ${thematics.join(', ')}`);

    for (
      let index = 0;
      index < thematics.length;
      index += 1
    ) {
      const input = await getThematicInput(index);

      await fillAutocomplete(
        input,
        thematics[index]
      );

      if (index < thematics.length - 1) {
        const addButton = await waitForElement(
          SELECTORS.thematicAddMore
        );

        await safeClick(
          addButton,
          'Ajouter un autre élément thématique'
        );

        await waitForAjax();
      }
    }

    return thematics;
  }

  /**
   * Sélectionne les publics/métiers dans Fancytree.
   *
   * Formats acceptés :
   *
   * "Tous les personnels"
   *
   * {
   *   path: "Parent > Enfant > Public"
   * }
   *
   * {
   *   label: "Tous les enseignants"
   * }
   */
  async function applyPublics({
    item = {},
    exportData = {},
  } = {}) {
    const exportPublics = Array.isArray(exportData.publics)
      ? exportData.publics
      : [];

    const itemPublics = Array.isArray(item.publics)
      ? item.publics
      : [];

    const publics = uniquePublics([
      ...exportPublics,
      ...itemPublics,
    ]);

    if (!publics.length) {
      await clickFancytreeTitle(
        'Tous les personnels',
        'Sélectionner le métier par défaut : Tous les personnels',
        {
          optional: true,
          scopeSelector: SELECTORS.publicTreeWrapper,
        }
      );

      return ['Tous les personnels'];
    }

    const selectedPaths = [];

    for (const publicEntry of publics) {
      const path = publicPath(publicEntry);

      if (!path.length) {
        continue;
      }

      const wasSelected = await clickFancytreePath(
        SELECTORS.publicTreeWrapper,
        path,
        `Sélectionner le métier : ${path.join(' > ')}`,
        {
          optional: true,
        }
      );

      if (wasSelected) {
        selectedPaths.push(path.join(' > '));
      }
    }

    return selectedPaths;
  }

  // ---------------------------------------------------------------------------
  // 3. LOGIQUE DES THÉMATIQUES
  // ---------------------------------------------------------------------------

  /**
   * Convertit les libellés contenus dans :
   *
   * item.theme
   * exportData.theme
   *
   * en identifiants Drupal.
   *
   * Ajoute ensuite les thématiques obligatoires.
   */
  function resolveThematics({
    item = {},
    exportData = {},
  } = {}) {
    const rawTheme = [
      item.theme,
      exportData.theme,
    ]
      .filter(Boolean)
      .join(', ');

    const normalizedTheme = normalize(rawTheme);
    const values = [];

    CONFIG.thematicMap.forEach(
      ([aliases, thematicIds]) => {
        const matches = aliases.some(alias => {
          return normalizedTheme.includes(
            normalize(alias)
          );
        });

        if (!matches) {
          return;
        }

        thematicIds.forEach(thematicId => {
          const value = String(thematicId).trim();

          if (
            value &&
            !values.includes(value)
          ) {
            values.push(value);
          }
        });
      }
    );

    CONFIG.requiredThematics.forEach(thematic => {
      const value = String(thematic).trim();

      if (
        value &&
        !values.includes(value)
      ) {
        values.push(value);
      }
    });

    return values;
  }

  /**
   * Récupère le champ thématique
   * correspondant à l’index demandé.
   */
  async function getThematicInput(index) {
    await waitForElement(
      SELECTORS.thematicInput
    );

    const inputs = [
      ...document.querySelectorAll(
        SELECTORS.thematicInput
      ),
    ].filter(isVisible);

    const input =
      inputs[index] ||
      inputs.find(candidate => !candidate.value) ||
      inputs[inputs.length - 1];

    if (!input) {
      throw new Error(
        `Champ thématique introuvable pour l’index ${index}.`
      );
    }

    return input;
  }

  /**
   * Saisit une valeur dans un autocomplete Drupal
   * et valide une suggestion.
   */
  async function fillAutocomplete(input, value) {
    if (!input) {
      throw new Error(
        `Champ d’autocomplétion introuvable pour : ${value}`
      );
    }

    input.focus();
    setValue(input, value);

    /*
     * Drupal/jQuery UI déclenche normalement
     * la recherche à partir de l’événement input.
     *
     * ArrowDown aide à ouvrir la liste
     * dans certaines configurations.
     */
    input.dispatchEvent(
      new KeyboardEvent('keydown', {
        key: 'ArrowDown',
        bubbles: true,
      })
    );

    await sleep(
      Math.max(
        CONFIG.delayMs,
        CONFIG.autocompleteDelayMs
      )
    );

    await waitUntil(
      () => visibleAutocompleteSuggestions().length > 0,
      5000
    ).catch(() => null);

    const suggestions =
      visibleAutocompleteSuggestions();

    const normalizedValue = normalize(value);

    const matchingSuggestion =
      suggestions.find(element => {
        const text = normalize(
          element.textContent
        );

        return (
          text.includes(normalizedValue) ||
          text.includes(`(${normalizedValue})`)
        );
      });

    const suggestion =
      matchingSuggestion ||
      suggestions[0];

    if (suggestion) {
      await safeClick(
        suggestion,
        `Valider l’autocomplétion : ${value}`
      );

      await waitUntil(
        () => {
          return (
            normalize(input.value) !== normalizedValue ||
            !isVisible(input)
          );
        },
        3000
      ).catch(() => null);
    } else {
      /*
       * Secours lorsque la liste jQuery UI
       * n’est pas détectée.
       */
      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'Enter',
          bubbles: true,
        })
      );

      input.blur();
    }

    await sleep(
      Math.max(
        CONFIG.delayMs,
        CONFIG.autocompleteSelectionDelayMs
      )
    );

    await waitForAjax();
  }

  function visibleAutocompleteSuggestions() {
    return [
      ...document.querySelectorAll(
        '.ui-autocomplete li, .ui-menu-item'
      ),
    ].filter(isVisible);
  }

  // ---------------------------------------------------------------------------
  // 4. LOGIQUE DES PUBLICS / MÉTIERS FANCYTREE
  // ---------------------------------------------------------------------------

  /**
   * Déroule les parents d’un chemin,
   * puis coche le dernier nœud.
   */
  async function clickFancytreePath(
    scopeSelector,
    path,
    message,
    options = {}
  ) {
    const cleanPath = path
      .map(String)
      .map(value => value.trim())
      .filter(Boolean);

    if (!cleanPath.length) {
      return false;
    }

    for (
      const title of cleanPath.slice(0, -1)
    ) {
      await ensureFancytreeExpanded(
        title,
        {
          ...options,
          scopeSelector,
        }
      );
    }

    return clickFancytreeTitle(
      cleanPath[cleanPath.length - 1],
      message,
      {
        ...options,
        scopeSelector,
      }
    );
  }

  /**
   * Déroule un nœud Fancytree
   * si nécessaire.
   */
  async function ensureFancytreeExpanded(
    title,
    options = {}
  ) {
    const node = findFancytreeNode(
      title,
      options
    );

    if (!node) {
      if (options.optional) {
        log(
          `Nœud public optionnel introuvable : ${title}`
        );

        return false;
      }

      throw new Error(
        `Nœud Fancytree introuvable : ${title}`
      );
    }

    const treeItem = node.closest(
      '[role="treeitem"]'
    );

    const isExpanded =
      treeItem?.getAttribute('aria-expanded') ===
      'true';

    if (!isExpanded) {
      const expander =
        node.querySelector(
          '.fancytree-expander'
        ) ||
        node
          .closest('.fancytree-node')
          ?.querySelector(
            '.fancytree-expander'
          );

      if (expander) {
        await safeClick(
          expander,
          `Dérouler ${title}`
        );

        await waitForAjax();
      }
    }

    return true;
  }

  /**
   * Coche la case associée
   * à un libellé Fancytree.
   */
  async function clickFancytreeTitle(
    title,
    message,
    options = {}
  ) {
    await waitUntil(
      () => {
        return (
          Boolean(
            findFancytreeNode(
              title,
              options
            )
          ) ||
          options.optional
        );
      },
      8000
    );

    const node = findFancytreeNode(
      title,
      options
    );

    if (!node) {
      log(
        `Champ public optionnel non trouvé : ${title}`
      );

      return false;
    }

    const checkbox =
      node.querySelector(
        '.fancytree-checkbox'
      ) ||
      node
        .closest('.fancytree-node')
        ?.querySelector(
          '.fancytree-checkbox'
        );

    if (!checkbox) {
      if (options.optional) {
        log(
          `Checkbox publique optionnelle introuvable : ${title}`
        );

        return false;
      }

      throw new Error(
        `Checkbox Fancytree introuvable : ${title}`
      );
    }

    /*
     * Évite de décocher un public
     * déjà sélectionné.
     */
    const treeItem = node.closest(
      '[role="treeitem"]'
    );

    const isSelected =
      treeItem?.getAttribute('aria-selected') ===
        'true' ||
      node.classList.contains(
        'fancytree-selected'
      ) ||
      node.classList.contains(
        'fancytree-partsel'
      );

    if (!isSelected) {
      await safeClick(
        checkbox,
        message
      );

      await waitForAjax();
    } else {
      log(
        `Public déjà sélectionné : ${title}`
      );
    }

    return true;
  }

  /**
   * Recherche un nœud à partir de son libellé,
   * avec tolérance sur :
   *
   * - les accents ;
   * - les tirets ;
   * - les espaces ;
   * - la casse.
   */
  function findFancytreeNode(
    title,
    options = {}
  ) {
    const wanted = normalize(title);

    const scope = options.scopeSelector
      ? document.querySelector(
          options.scopeSelector
        )
      : document;

    if (!scope) {
      return null;
    }

    return [
      ...scope.querySelectorAll(
        '.fancytree-node'
      ),
    ].find(node => {
      const text = normalize(
        node.querySelector(
          '.fancytree-title'
        )?.textContent || ''
      );

      return options.contains
        ? text.includes(wanted)
        : text === wanted;
    });
  }

  /**
   * Déduplique les publics sans transformer
   * les objets en "[object Object]".
   */
  function uniquePublics(values) {
    const seen = new Set();

    return values.filter(value => {
      if (!value) {
        return false;
      }

      const key =
        typeof value === 'string'
          ? normalize(value)
          : normalize(
              value.id ||
              value.path ||
              value.label
            );

      if (
        !key ||
        seen.has(key)
      ) {
        return false;
      }

      seen.add(key);
      return true;
    });
  }

  /**
   * Transforme un public exporté
   * en chemin Fancytree.
   */
  function publicPath(value) {
    if (typeof value === 'string') {
      return [
        value.trim(),
      ].filter(Boolean);
    }

    if (value?.path) {
      return String(value.path)
        .split('>')
        .map(part => part.trim())
        .filter(Boolean);
    }

    return [
      value?.label ||
      value?.id,
    ]
      .map(entry => {
        return String(
          entry || ''
        ).trim();
      })
      .filter(Boolean);
  }

  // ---------------------------------------------------------------------------
  // 5. DÉPENDANCES TECHNIQUES COMMUNES
  // ---------------------------------------------------------------------------

  async function safeClick(
    element,
    message
  ) {
    if (!element) {
      throw new Error(
        `Élément introuvable pour : ${message}`
      );
    }

    element.scrollIntoView({
      block: 'center',
      inline: 'center',
    });

    element.classList.add(
      'eafc-auto-highlight'
    );

    await sleep(120);

    if (
      CONFIG.safeMode &&
      !window.confirm(`${message} ?`)
    ) {
      element.classList.remove(
        'eafc-auto-highlight'
      );

      throw new Error(
        `Action annulée : ${message}`
      );
    }

    triggerClick(element);

    element.classList.remove(
      'eafc-auto-highlight'
    );

    await sleep(CONFIG.delayMs);
  }

  function triggerClick(element) {
    element.focus?.({
      preventScroll: true,
    });

    [
      'pointerover',
      'pointerenter',
      'mouseover',
      'mouseenter',
      'pointerdown',
      'mousedown',
      'pointerup',
      'mouseup',
    ].forEach(type => {
      element.dispatchEvent(
        new MouseEvent(type, {
          bubbles: true,
          cancelable: true,
        })
      );
    });

    element.click();
  }

  function setValue(
    element,
    value
  ) {
    if (!element) {
      return;
    }

    element.focus();

    element.value =
      value == null
        ? ''
        : String(value);

    element.dispatchEvent(
      new Event('input', {
        bubbles: true,
      })
    );

    element.dispatchEvent(
      new Event('change', {
        bubbles: true,
      })
    );
  }

  async function waitForElement(
    selector,
    timeout = 10000
  ) {
    await waitUntil(
      () => findVisibleElement(selector),
      timeout
    );

    return findVisibleElement(selector);
  }

  function findVisibleElement(selector) {
    return [
      ...document.querySelectorAll(selector),
    ].find(isVisible);
  }

  async function waitUntil(
    predicate,
    timeout = 10000
  ) {
    const start = Date.now();

    while (
      Date.now() - start < timeout
    ) {
      const result = predicate();

      if (result) {
        return result;
      }

      await sleep(150);
    }

    throw new Error(
      'Délai dépassé en attendant un élément ou une condition.'
    );
  }

  async function waitForAjax() {
    await sleep(CONFIG.delayMs);

    await waitUntil(
      () => {
        return !document.querySelector(
          '.ajax-progress, ' +
          '.ui-dialog .ajax-progress'
        );
      },
      15000
    ).catch(() => null);
  }

  /**
   * Normalisation tolérante :
   *
   * - suppression des accents ;
   * - homogénéisation des tirets ;
   * - suppression des espaces multiples ;
   * - passage en minuscules.
   */
  function normalize(value) {
    return String(value || '')
      .normalize('NFD')
      .replace(
        /[\u0300-\u036f]/g,
        ''
      )
      .replace(
        /[–—−]/g,
        '-'
      )
      .replace(
        /\s+/g,
        ' '
      )
      .trim()
      .toLowerCase();
  }

  function isVisible(element) {
    if (!element) {
      return false;
    }

    const style =
      window.getComputedStyle(element);

    return (
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      element.getClientRects().length > 0
    );
  }

  function sleep(ms) {
    return new Promise(resolve => {
      window.setTimeout(
        resolve,
        ms
      );
    });
  }

  function log(message) {
    console.log(
      `[EAFC publics/thématiques] ${message}`
    );
  }

  // ---------------------------------------------------------------------------
  // 6. EXPOSITION POUR CODEX OU LE SCRIPT PRINCIPAL
  // ---------------------------------------------------------------------------

  window.eafcPublicThematicTools = {
    applyAll,
    applyThematics,
    applyPublics,
    resolveThematics,
    publicPath,
    uniquePublics,

    /**
     * Permet de modifier la configuration
     * après le chargement du module.
     */
    configure(options = {}) {
      if (
        Number.isFinite(
          Number(options.delayMs)
        )
      ) {
        CONFIG.delayMs = Math.max(
          200,
          Number(options.delayMs)
        );
      }

      if (
        typeof options.safeMode ===
        'boolean'
      ) {
        CONFIG.safeMode =
          options.safeMode;
      }

      if (
        Array.isArray(
          options.requiredThematics
        )
      ) {
        CONFIG.requiredThematics = [
          ...options.requiredThematics,
        ];
      }

      if (
        Array.isArray(
          options.thematicMap
        )
      ) {
        CONFIG.thematicMap = [
          ...options.thematicMap,
        ];
      }

      return {
        ...CONFIG,
      };
    },

    /**
     * Retourne la configuration actuelle.
     */
    getConfig() {
      return {
        ...CONFIG,

        requiredThematics: [
          ...CONFIG.requiredThematics,
        ],

        thematicMap: [
          ...CONFIG.thematicMap,
        ],
      };
    },
  };

  // Ajout du style de surbrillance
  // si le module est utilisé seul.
  const style =
    document.createElement('style');

  style.textContent = `
    .eafc-auto-highlight {
      outline: 4px solid #ff6f4c !important;
      outline-offset: 3px !important;
    }
  `;

  document.head.appendChild(style);

  log(
    'Module chargé : window.eafcPublicThematicTools'
  );
})();