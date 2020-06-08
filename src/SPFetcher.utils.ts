import { sp, IList, IField } from '@pnp/sp/presets/all';
import {
  SPHttpClient,
  ISPHttpClientOptions,
  SPHttpClientConfiguration
} from '@microsoft/sp-http';
import {
  IListField,
  FieldLookup,
  ITerm,
  SPFetcherStructure
} from './interfaces';
import { SPFetcherInitializer } from './SPFetcher.initializer';

export class SPFetcherUtils<
  T extends SPFetcherStructure
> extends SPFetcherInitializer<T> {
  /**
   * Utility method: Initialize empty batch object
   */
  public createBatch() {
    return this.ready().then(() => sp.createBatch());
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
  public getListByTitle(
    title: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web => web.lists.getByTitle(title));
  }

  /**
   * Utility method: Get list by id
   */
  public getListById(
    id: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web => web.lists.getById(id));
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
  public getFieldsByListId(
    id: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getListById(id, site).then(list => this.getListFields(list));
  }

  /**
   * Utility method: Get all fields in list by name
   */
  public getFieldsByListTitle(
    title: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getListByTitle(title, site).then(list =>
      this.getListFields(list)
    );
  }

  /**
   * Utility method: Get reference to field by id
   */
  public getFieldById(
    id: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web => web.fields.getById(id));
  }

  /**
   * Utility method: Get reference to field by title
   */
  public getFieldByTitle(
    title: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web => web.fields.getByTitle(title));
  }

  /**
   * Utility method: Get reference to field by internal name or title
   */
  public getFieldByInternalNameOrTitle(
    title: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web =>
      web.fields.getByInternalNameOrTitle(title)
    );
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
  public getLookupByFieldId(
    id: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getFieldById(id, site).then(field =>
      this.getFieldLookup(field)
    );
  }

  /**
   * Utility method: Get reference to relevant lookup list by field title
   */
  public getLookupByFieldTitle(
    title: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getFieldByTitle(title, site).then(field =>
      this.getFieldLookup(field)
    );
  }

  /**
   * Utility method: Get all fields of a taxonomy termset
   */
  public getTermsetById(
    id: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.post(
      `${site || this.sites.current}/_vti_bin/client.svc/ProcessQuery`,
      {
        headers: {
          accept: 'application/json',
          'content-type': 'application/xml'
        },
        body: `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Javascript Library"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectIdentityQuery Id="2" ObjectPathId="0" /><ObjectPath Id="4" ObjectPathId="3" /><ObjectIdentityQuery Id="5" ObjectPathId="3" /><ObjectPath Id="7" ObjectPathId="6" /><ObjectIdentityQuery Id="8" ObjectPathId="6" /><ObjectPath Id="10" ObjectPathId="9" /><Query Id="11" ObjectPathId="6"><Query SelectAllProperties="true"><Properties /></Query></Query><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties /></Query><ChildItemQuery SelectAllProperties="false"><Properties><Property Name="IsRoot" SelectAll="true" /><Property Name="Labels" SelectAll="true" /><Property Name="TermsCount" SelectAll="true" /><Property Name="CustomSortOrder" SelectAll="true" /><Property Name="Id" SelectAll="true" /><Property Name="Name" SelectAll="true" /><Property Name="PathOfTerm" SelectAll="true" /><Property Name="Parent" SelectAll="true" /><Property Name="LocalCustomProperties" SelectAll="true" /><Property Name="IsDeprecated" ScalarProperty="true" /><Property Name="IsAvailableForTagging" ScalarProperty="true" /></Properties></ChildItemQuery></Query></Actions><ObjectPaths><StaticMethod Id="0" Name="GetTaxonomySession" TypeId="{981cbc68-9edc-4f8d-872f-71146fcbb84f}" /><Method Id="3" ParentId="0" Name="GetDefaultKeywordsTermStore" /><Method Id="6" ParentId="3" Name="GetTermSet"><Parameters><Parameter Type="Guid">${id}</Parameter></Parameters></Method><Method Id="9" ParentId="6" Name="GetAllTerms" /></ObjectPaths></Request>`
      }
    )
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
   * Utility method: Get default document library id
   * No need to use ready() here because getProperties() takes care of that.
   */
  public getDefaultLibraryId(site?: keyof SPFetcherInitializer<T>['sites']) {
    return this.Web(site)
      .then(web => web.defaultDocumentLibrary.select('Id').get())
      .then(r => r.Id)
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
  public getDefaultLibrary(site?: keyof SPFetcherInitializer<T>['sites']) {
    return this.getDefaultLibraryId(site).then(libraryId =>
      this.getListById(libraryId, site)
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
  public getItemByPath(
    path: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.Web(site).then(web =>
      web
        .getFileByServerRelativePath(path)
        .getItem()
        .then(item => item)
    );
  }

  /**
   * Utility method: Auto get parent library
   */
  public getParentLibrary(
    path?: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return path
      ? this.getItemByPath(path, site)
          .then(item => item.toUrl().replace(/^.*guid'(.*)'(.*)/g, '$1'))
          .then(libraryId => this.web.lists.getById(libraryId))
      : this.getDefaultLibrary(site);
  }

  /**
   * Utility method: Get all items of parent
   */
  public getAllItems(
    parent?: string,
    contentType?: string,
    select?: string | string[],
    filter?: string | string[],
    top?: number,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    if (parent) parent = parent.replace(/^\/|\/$/g, '');
    const filters = [
      parent ? `substringof('${parent}/',FileRef)` : undefined,
      contentType ? `startswith(ContentTypeId,'${contentType}')` : undefined
    ]
      .concat(filter || [])
      .filter(test => test)
      .join(' and ');
    return this.getParentLibrary(parent, site)
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
    top?: number,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getAllItems(
      parent,
      '0x0101',
      select,
      filter,
      top,
      site
    ).then(items => items.get());
  }

  /**
   * Utility method: Get all folders
   */
  public getAllFolders(
    parent?: string,
    select?: string | string[],
    filter?: string | string[],
    top?: number,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getAllItems(
      parent,
      '0x0120',
      select,
      filter,
      top,
      site
    ).then(items => items.get());
  }
}
