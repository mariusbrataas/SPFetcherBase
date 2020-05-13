import { sp } from '@pnp/sp';
import { Web, IWeb } from '@pnp/sp/webs';
import '@pnp/sp/lists/web';
import '@pnp/sp/files/web';
import '@pnp/sp/folders/web';
import '@pnp/sp/fields/web';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import { IListField } from './interfaces';
import { IList } from '@pnp/sp/lists';

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
    base: string;
  };
  protected lists: {};
  protected termsets: {};

  protected web: IWeb;

  public status: 'not initialized' | 'initializing' | 'ready' | 'error';
  protected queue: (() => void)[];

  // Constructor
  constructor() {
    this.urls = {
      absolute: undefined,
      base: undefined
    };
    this.lists = {};
    this.termsets = {};
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
   */
  public initialize(context: BaseComponentContext): Promise<void> {
    return new Promise(resolve => {
      if (this.status === 'initializing') {
        this.queue.push(resolve);
      } else {
        this.status = 'initializing';
        sp.setup({ spfxContext: context });
        this.context = context;
        this.urls.absolute = this.context.pageContext.site.absoluteUrl;
        this.urls.base = this.urls.absolute.match(/(.*).sharepoint.com/)[0];
        this.getWeb()
          .then(web => (this.web = web))
          .then(() => resolve());
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
      } else this.queue.push(resolve);
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
   * Get a new web object.
   * If the argument is "base" the base url will be used.
   *
   * @param base
   */
  public getWeb(base?: string) {
    return this.ready().then(() =>
      base
        ? Web(base === 'base' ? this.urls.base : base)
        : this.web || Web(this.urls.absolute)
    );
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
   * Utility method: Get storage entity
   * @param entity
   */
  public getStorageEntity(entity: string) {
    return this.ready().then(() => this.web.getStorageEntity(entity));
  }

  /**
   * Utility method: Get list by title
   */
  public getListByTitle(title: string) {
    return this.ready().then(() => this.web.lists.getByTitle(title));
  }

  /**
   * Utility method: Get list by id
   */
  public getListById(id: string) {
    return this.ready().then(() => this.web.lists.getById(id));
  }

  /**
   * Utility method: Get fields of a list
   */
  public getListFields(list: IList): Promise<IListField[]> {
    return list.relatedFields.get();
  }

  /**
   * Utility method: Get all fields in list by id
   */
  public getFieldsByListId(...ids: string[]) {
    return Promise.all(
      ids.map(id => this.getListById(id).then(list => this.getListFields(list)))
    ).then(r => r.reduce((prev, fields) => prev.concat(fields), []));
  }

  /**
   * Utility method: Get all fields in list by name
   */
  public getFieldsByListTitle(...titles: string[]) {
    return Promise.all(
      titles.map(title =>
        this.getListByTitle(title).then(list => this.getListFields(list))
      )
    ).then(r => r.reduce((prev, fields) => prev.concat(fields), []));
  }

  /**
   * Utility method: Get interface for list
   */
  public getFieldsInterface(list: IList) {
    return Promise.all([
      list.get().then(r => (r.Title || 'List').split(' ').join('')),
      this.getListFields(list)
        .then(fields =>
          fields.reduce(
            (prev, field) => ({
              ...prev,
              [field.InternalName]: field['odata.type'].split('.').slice(-1)[0]
            }),
            {}
          )
        )
        .then(r =>
          Object.keys(r)
            .sort()
            .map(key => `\n  ${key}: ${r[key]};`)
            .join('')
        )
    ]).then(([name, types]) => `interface I${name}Item {${types}\n}`);
  }

  public getFieldsInterfaceByListId(...ids: string[]) {
    return Promise.all(
      ids.map(id =>
        this.getListById(id).then(list => this.getFieldsInterface(list))
      )
    ).then(r => r.join('\n\n'));
  }

  public getFieldsInterfaceByListTitle(...titles: string[]) {
    return Promise.all(
      titles.map(title =>
        this.getListByTitle(title).then(list => this.getFieldsInterface(list))
      )
    ).then(r => r.join('\n\n'));
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
   * Utility method: Get current document library id
   */
  public getCurrentLibraryId() {
    return this.ready().then(() => this.context.pageContext.list.id.toString());
  }

  /**
   * Utility method: Get current document library
   */
  public getCurrentLibrary() {
    return this.getCurrentLibraryId().then(libraryId =>
      this.getListById(libraryId)
    );
  }

  /**
   * Utility method: Check whether the user is curently viewing the default library.
   */
  public isDefaultLibrary() {
    return Promise.all([
      this.getDefaultLibraryId(),
      this.getCurrentLibraryId()
    ]).then(([defaultId, currentId]) => defaultId === currentId);
  }

  /**
   * Utility method: Get item by path
   */
  public getItemByPath(path: string) {
    return this.ready().then(() =>
      this.web
        .getFileByServerRelativePath(path)
        .getItem()
        .then(item => item)
    );
  }

  /**
   * Utility method: Auto get parent library
   */
  public getParentLibrary(path?: string) {
    return path
      ? this.getItemByPath(path)
          .then(item => item.toUrl().replace(/^.*guid'(.*)'(.*)/g, '$1'))
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
