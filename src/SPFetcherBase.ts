import { sp, Web, IWeb, IList, IField } from '@pnp/sp/presets/all';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientConfiguration
} from '@microsoft/sp-http';
import { IListField, FieldLookup, ITerm, TaxonomyField } from './interfaces';

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
   * Utility method: Perform a fetch-request using the spHttpClient
   */
  public fetch(
    url: string,
    options?: ISPHttpClientOptions,
    config: SPHttpClientConfiguration = SPHttpClient.configurations.v1,
    method: 'get' | 'post' = 'get'
  ) {
    return this.ready().then(() =>
      this.context.spHttpClient[method](
        url.startsWith('https://')
          ? url
          : `${this.urls.base}/${url.replace(/^\/+/g, '')}`,
        config,
        options
      )
    );
  }

  /**
   * Utility method: Perform a get-request using the spHttpClient
   */
  public get(
    url: string,
    options?: ISPHttpClientOptions,
    config: SPHttpClientConfiguration = SPHttpClient.configurations.v1
  ) {
    return this.fetch(url, options, config, 'get');
  }

  /**
   * Utility method: Perform a post-request using the spHttpClient
   */
  public post(
    url: string,
    options?: ISPHttpClientOptions,
    config: SPHttpClientConfiguration = SPHttpClient.configurations.v1
  ) {
    return this.fetch(url, options, config, 'post');
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
    return list.fields();
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
   * Utility method: Get reference to field by id
   */
  public getFieldById(id: string) {
    return this.ready().then(() => this.web.fields.getById(id));
  }

  /**
   * Utility method: Get reference to field by title
   */
  public getFieldByTitle(title: string) {
    return this.ready().then(() => this.web.fields.getByTitle(title));
  }

  /**
   * Utility method: Get reference to the relevant field's lookup list
   */
  public getFieldLookup(field: IField) {
    return field.get().then((r: FieldLookup) => this.getListById(r.LookupList));
  }

  /**
   * Utility method: Get reference to relevant lookup list by field id
   */
  public getLookupByFieldId(id: string) {
    return this.getFieldById(id).then(field => this.getFieldLookup(field));
  }

  /**
   * Utility method: Get reference to relevant lookup list by field title
   */
  public getLookupByFieldTitle(title: string) {
    return this.getFieldByTitle(title).then(field =>
      this.getFieldLookup(field)
    );
  }

  /**
   * Utility method: Get all fields of a taxonomy termset
   */
  public getTermsetById(id: string) {
    return this.post(`${this.urls.absolute}/_vti_bin/client.svc/ProcessQuery`, {
      headers: {
        accept: 'application/json',
        'content-type': 'application/xml'
      },
      body: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectIdentityQuery Id="2" ObjectPathId="0" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="IsRoot" SelectAll="true" /><Property Name="Labels" SelectAll="true" /><Property Name="TermsCount" SelectAll="true" /><Property Name="CustomSortOrder" SelectAll="true" /><Property Name="Id" SelectAll="true" /><Property Name="Name" SelectAll="true" /><Property Name="PathOfTerm" SelectAll="true" /><Property Name="Parent" SelectAll="true" /><Property Name="LocalCustomProperties" SelectAll="true" /><Property Name="IsDeprecated" ScalarProperty="true" /><Property Name="IsAvailableForTagging" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="0" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="3" ParentId="0" Name="GetDefaultKeywordsTermStore" /><Method Id="6" ParentId="3" Name="GetTermSet"><Parameters><Parameter Type="Guid">${id}</Parameter></Parameters></Method><Method Id="9" ParentId="6" Name="GetAllTerms" /></ObjectPaths></Request>`
    })
      .then(r => r.json())
      .then(
        r =>
          r.filter(
            (test: any) => test._ObjectType_ === 'SP.Taxonomy.TermCollection'
          )[0]._Child_Items_ as ITerm[]
      )
      .then(r => {
        r.forEach(term => {
          term.Id = term.Id.replace(/^.*Guid\((.*)\)(.*)/g, '$1');
          if (term.Parent)
            term.Parent.Id = term.Parent.Id.replace(
              /^.*Guid\((.*)\)(.*)/g,
              '$1'
            );
        });
        return r;
      });
  }

  /**
   * Helper method: Get the taxonomy node closest to the requested path
   * @param termset_id
   * @param new_path
   */
  private getTaxonomyClosestParent(termset_id: string, new_path: string) {
    const split_path = new_path.split(';');
    return this.getTermsetById(termset_id).then(
      r =>
        r.reduce(
          ([highscore, highterm], term) => {
            const score = split_path.reduce(
              (score, _, idx) =>
                split_path.slice(0, idx + 1).join(';') === term.PathOfTerm
                  ? idx
                  : score,
              -1
            );
            return score > highscore ? [score, term] : [highscore, highterm];
          },
          [-1, undefined]
        ) as [number, ITerm]
    );
  }

  /**
   * Helper method: Add a new node to taxonomy
   * @param sspid
   * @param termset_id
   * @param new_name
   * @param parent_term
   */
  private addTaxonomyItem(
    sspid: string,
    termset_id: string,
    new_name: string,
    parent_term: Pick<ITerm, 'Id' | 'IsRoot'>
  ) {
    return this.post(
      '_vti_bin/taxonomyinternalservice.json/CreateTaxonomyItem',
      {
        body: JSON.stringify({
          sspId: sspid,
          lcid: 1033,
          parentType: parent_term.IsRoot ? 3 : 4,
          webId: '00000000-0000-0000-0000-000000000000',
          listId: '00000000-0000-0000-0000-000000000000',
          parentId: parent_term.Id,
          termsetId: termset_id,
          newName: new_name
        })
      }
    );
  }

  /**
   * Helper method: Build entire taxonomy path
   * @param sspid
   * @param termset_id
   * @param new_path
   * @param prev_name
   */
  public buildTaxonomyPath(
    sspid: string,
    termset_id: string,
    new_path: string,
    prev_name: string = ''
  ): Promise<boolean> {
    const path = new_path;
    const split_path = path.split(';');
    return this.getTaxonomyClosestParent(termset_id, new_path).then(
      ([highscore, term]) => {
        const new_name = split_path[highscore + 1];
        if (prev_name === new_name)
          throw new Error(
            `Failed to add '${new_name}' to termset. termsetId: ${termset_id}, sspid: ${sspid}`
          );
        return new_name
          ? new_name === prev_name
            ? null
            : this.addTaxonomyItem(
                sspid,
                termset_id,
                new_name,
                term
                  ? { ...term, IsRoot: false }
                  : { Id: termset_id, IsRoot: true }
              ).then(() =>
                this.buildTaxonomyPath(sspid, termset_id, new_path, new_name)
              )
          : true;
      }
    );
  }

  /**
   * Utility method: Build taxonomy path at termset related to field
   * @param field
   * @param new_path
   */
  public buildTaxonomyPathByField(field: IField, new_path: string) {
    return field
      .get()
      .then((field: TaxonomyField) =>
        this.buildTaxonomyPath(field.SspId, field.TermSetId, new_path)
      );
  }

  /**
   * Utility method: Build taxonomy path at termset related to field by field id
   * @param id
   * @param new_path
   */
  public buildTaxonomyPathByFieldId(id: string, new_path: string) {
    return this.getFieldById(id).then(field =>
      this.buildTaxonomyPathByField(field, new_path)
    );
  }

  /**
   * Utility method: Build taxonomy path at termset related to field by field title
   * @param title
   * @param new_path
   */
  public buildTaxonomyPathByFieldTitle(title: string, new_path: string) {
    return this.getFieldByTitle(title).then(field =>
      this.buildTaxonomyPathByField(field, new_path)
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
