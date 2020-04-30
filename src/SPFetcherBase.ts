import pnp, { Web } from 'sp-pnp-js';
import { BaseComponentContext } from '@microsoft/sp-component-base';

/**
 * SPFetcher base
 * A container for all requests in a project.
 *
 * Start by extending this class and create your own constructor and startup routines.
 *
 * Usage:
 * - Put startup-routines in the method "startupRoutines". This will be called automatically during initialization.
 * - Start your own methods with return this.ready().then(() => { ... }). This ensures execution once ready.
 * - While loading your webpart, call Fetcher.initialize to set it up
 *
 * @example
 * class FetcherClass extends SPFetcherBase {
 *   // Properties
 *   protected urls: {
 *     absolute: string;
 *     logic: string;
 *   };
 *
 *   // Constructor
 *   constructor() {
 *     super();
 *     this.urls = {
 *       absolute: undefined,
 *       logic: undefined,
 *     };
 *   }
 *
 *   // Startup routines
 *   protected startupRoutines(): Promise<any> {
 *     return Promise.all([
 *       this.web
 *         .getStorageEntity("SettingsSite")
 *         .then((r: any) =>
 *           new Web(r.Value).lists
 *             .getByTitle("Settings")
 *             .items.filter("Title eq 'CopyFileLogicAppUrl'")
 *             .get()
 *         )
 *         .then((r) => (this.urls.logic = r[0].value)),
 *     ]);
 *   }
 *
 *   // Example method: get
 *   public getFiles() {
 *     return this.getDefaultLibrary().then((library) =>
 *       library.items
 *         .filter(
 *           `startswith(ContentTypeId,'0x0101')`
 *         )
 *         .get()
 *     );
 *   }
 * }
 */
export class SPFetcherBase {
  // Properties
  public context: BaseComponentContext;
  protected urls: {
    absolute: string;
  };
  protected web: Web;

  public log: Array<{ time: string; title: string }>;

  public status: 'not initialized' | 'initializing' | 'ready' | 'error';
  protected queue: (() => void)[];

  // Constructor
  constructor() {
    this.urls = {
      absolute: undefined
    };
    this.status = 'not initialized';
    this.queue = [];
  }

  /**
   * Main initializer
   * Accepts the webpart's context and sets up pnp.
   *
   * Put other startup routines in the startupRoutines method. It will be called during init.
   *
   * @param context
   * @param logging
   */
  public initialize(
    context: BaseComponentContext,
    logging?: boolean
  ): Promise<void> {
    return new Promise(resolve => {
      if (this.status === 'initializing') {
        this.queue.push(resolve);
      } else {
        this.status = 'initializing';
        pnp.setup({ spfxContext: context });
        this.context = context;
        this.urls.absolute = this.context.pageContext.site.absoluteUrl;
        this.web = new Web(this.urls.absolute);
        if (logging) this.initLogging();
        resolve();
      }
    })
      .then(() => this.startupRoutines())
      .then(() => (this.status = 'ready'))
      .then(() => {
        this.queue.forEach(callback => callback());
        this.queue = [];
      })
      .catch((error: Error) => {
        this.status = 'error';
        throw error;
      });
  }

  /**
   * Add a new entry to the log
   *
   * @param title
   */
  private addLog(title: string) {
    if (this.log)
      this.log.push({
        title,
        time: new Date().toISOString().replace(/(T(.*)Z)|[^]/g, '$2')
      });
    return true;
  }

  /**
   * Initialize logging
   *
   * Adds a "middleman" to all class methods, ensuring a new entry with a
   * timestamp and the title of the method is added to the log whenever the
   * method is called
   */
  private initLogging() {
    if (this.log === undefined) {
      this.log = [];
      Object.getOwnPropertyNames(Object.getPrototypeOf(this))
        .filter(test =>
          [
            'constructor',
            'initLogging',
            'initialize',
            'addLog',
            'startupRoutines',
            'ready'
          ].every(name => name !== test)
        )
        .forEach(name => {
          const func = this[name];
          this[name] = (...args: any[]) => this.addLog(name) && func(...args);
        });
    }
  }

  /**
   * Execute immediately if ready.
   * Otherwise add promise to queue. It will be resolved during initialization.
   *
   * A tip: Always use this one to initialize a promise chain.
   *
   * @example
   * public myFunction() {
   *   return this.ready().then(() => {
   *     // your code here
   *   });
   * }
   */
  protected ready(): Promise<void> {
    return new Promise(resolve => {
      if (this.status === 'ready' || this.status === 'initializing') {
        resolve();
      } else this.queue.push(() => this.ready().then(resolve));
    });
  }

  /**
   * Custom startup routines.
   * Put all your startup routines here.
   * This method will be called by the initializer after basic setup, but before
   * tasks in the queue are executed.
   */
  protected startupRoutines(): Promise<any> {
    return Promise.all([]);
  }

  /**
   * Utility method: Get all properties
   */
  public getProperties(): Promise<any> {
    return this.ready().then(() =>
      this.web.select('AllProperties').expand('AllProperties').get()
    );
  }

  /**
   * Utility method: Get default document library id
   * No need to use ready() here because getProperties() takes care of that.
   */
  public getDefaultLibraryId() {
    return this.getProperties()
      .then(
        r => r.AllProperties && (r.AllProperties.GroupDocumentsListId as string)
      )
      .then(libraryId => {
        if (libraryId) {
          return libraryId;
        } else throw new Error('Could not find default documents library.');
      });
  }

  /**
   * Utility method: Get default documents library
   * No need to use ready() here because getProperties() takes care of that.
   */
  public getDefaultLibrary() {
    return this.getDefaultLibraryId().then(libraryId =>
      this.web.lists.getById(libraryId)
    );
  }

  /**
   * Utility method: Check whether the user is curently viewing the default library.
   */
  public isDefaultLibrary() {
    return this.getDefaultLibraryId().then(
      libraryId => libraryId === this.context.pageContext.list.id.toString()
    );
  }

  /**
   * Utility method: Auto get parent library
   */
  public getParentLibrary(parent?: string) {
    if (parent) parent = parent.replace(/^\/|\/$/g, '');
    return parent
      ? this.web
          .getFolderByServerRelativePath(`/${parent}`)
          .getItem()
          .then(item =>
            item
              .toUrl()
              .match(/guid'(.*)Items/g)[0]
              .match(/'(.*)'/g)[0]
              .slice(1, -1)
          )
          .then(libraryId => this.web.lists.getById(libraryId))
      : this.getDefaultLibrary();
  }

  /**
   * Utility method: Get all items of parent
   */
  public getAllItems(
    parent?: string,
    contentType?: string,
    select?: string | string[],
    filter?: string | string[],
    top?: number
  ) {
    if (parent) parent = parent.replace(/^\/|\/$/g, '');
    const filters = [
      parent ? `substringof('${parent}/',FileRef)` : undefined,
      contentType ? `startswith(ContentTypeId,'${contentType}')` : undefined
    ]
      .concat(filter || [])
      .filter(test => test)
      .join(' and ');
    return this.getParentLibrary(parent)
      .then(library => library.items)
      .then(items =>
        filters && filters.length ? items.filter(filters) : items
      )
      .then(items => (select ? items.select(...[].concat(select)) : items))
      .then(items => (top ? items.top(top) : items));
  }

  /**
   * Utility method: Get all files
   */
  public getAllFiles(
    parent?: string,
    select?: string | string[],
    filter?: string | string[],
    top?: number
  ) {
    return this.getAllItems(parent, '0x0101', select, filter, top).then(items =>
      items.get()
    );
  }

  /**
   * Utility method: Get all folders
   */
  public getAllFolders(
    parent?: string,
    select?: string | string[],
    filter?: string | string[],
    top?: number
  ) {
    return this.getAllItems(parent, '0x0120', select, filter, top).then(items =>
      items.get()
    );
  }
}
