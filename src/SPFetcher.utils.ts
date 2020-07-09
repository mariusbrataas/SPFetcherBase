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
  SPFetcherStructure,
  ItemType
} from './interfaces';
import { SPFetcherInitializer } from './SPFetcher.initializer';

export class SPFetcherUtils<
  T extends SPFetcherStructure
> extends SPFetcherInitializer<T> {
  /**
   * Properties
   */
  private contentTypes: {
    [key: string]: {
      callbacks: ((arg: string[]) => void)[];
      StringIds: string[];
    };
  };

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
   * Utility method: Search for users
   */
  public searchUsers(query: string, limit: number = 5) {
    return this.ready().then(() =>
      sp.profiles.clientPeoplePickerSearchUser({
        AllowEmailAddresses: true,
        AllowMultipleEntities: false,
        AllUrlZones: false,
        MaximumEntitySuggestions: limit,
        PrincipalType: 1,
        QueryString: query
      })
    );
  }

  /**
   * Utility method: Get all properties
   */
  public getProperties(
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ): Promise<any> {
    return this.Web(site).then(web =>
      web.select('AllProperties').expand('AllProperties').get()
    );
  }

  /**
   * Utility method: Get storage entity
   * @param entity
   */
  public getStorageEntity(
    entity: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.Web(site).then(web => web.getStorageEntity(entity));
  }

  /**
   * Utility method: Get list by title
   */
  public getListByTitle(
    title: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.Web(site).then(web => web.lists.getByTitle(title));
  }

  /**
   * Utility method: Get list by id
   */
  public getListById(
    id: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.getListById(id, site).then(list => this.getListFields(list));
  }

  /**
   * Utility method: Get all fields in list by name
   */
  public getFieldsByListTitle(
    title: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.Web(site).then(web => web.fields.getById(id));
  }

  /**
   * Utility method: Get reference to field by title
   */
  public getFieldByTitle(
    title: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.Web(site).then(web => web.fields.getByTitle(title));
  }

  /**
   * Utility method: Get reference to field by internal name or title
   */
  public getFieldByInternalNameOrTitle(
    title: string,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.post(
      `${
        this.sites[site] || this.sites.current
      }/_vti_bin/client.svc/ProcessQuery`,
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
  public getDefaultLibraryId(
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
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
  public getDefaultLibrary(
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
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
    type?: ItemType,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.Web(site).then(web =>
      web[
        type === 'folder'
          ? 'getFolderByServerRelativePath'
          : 'getFileByServerRelativePath'
      ](`/${path.replace(/^\/|\/$/g, '')}`)
        .getItem()
        .then(item => item)
    );
  }

  /**
   * Utility method: Auto get parent library
   */
  public getParentLibrary(
    path?: string,
    type?: ItemType,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return path
      ? this.getItemByPath(path, type, site)
          .then(item => item.toUrl().replace(/^.*guid'(.*)'(.*)/g, '$1'))
          .then(libraryId => this.getListById(libraryId, site))
      : this.getDefaultLibrary(site);
  }

  /**
   * Utility method: Get all items of parent
   */
  public getAllItems(
    parent?: string,
    type?: ItemType,
    select?: string | string[],
    filter?: string | string[],
    top?: number,
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    if (parent) parent = parent.replace(/^\/|\/$/g, '');
    const filters = [
      parent
        ? `FileRef eq '${parent}' or substringof('${parent}/',FileRef)`
        : undefined
    ]
      .concat(filter || [])
      .filter(test => test)
      .join(' and ');
    return this.getParentLibrary(parent, type, site)
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.getAllItems(
      parent,
      'file',
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
    site?: Parameters<SPFetcherInitializer<T>['Web']>[0]
  ) {
    return this.getAllItems(
      parent,
      'folder',
      select,
      filter,
      top,
      site
    ).then(items => items.get());
  }

  /**
   * Helper method: Get ids of all content types
   */
  private getContentTypeIds({
    site,
    list
  }: {
    site?: string;
    list?: string;
  }): Promise<string[]> {
    const key = `[${site || 'site'}]-[${list || 'list'}]`;
    if (!this.contentTypes) this.contentTypes = {};
    return new Promise(resolve => {
      if (this.contentTypes[key]) {
        if (this.contentTypes[key].StringIds) {
          resolve(this.contentTypes[key].StringIds);
        } else this.contentTypes[key].callbacks.push(resolve);
      } else {
        this.contentTypes[key] = {
          callbacks: [resolve],
          StringIds: undefined
        };
        (list === undefined
          ? this.Web(site).then(web => web.contentTypes)
          : this.getListByTitle(list, site).then(
              library => library.contentTypes
            )
        )
          .then(contentTypes => contentTypes.select('StringId').get())
          .then(contentTypes => contentTypes.map(({ StringId }) => StringId))
          .then(StringIds => {
            this.contentTypes[key].StringIds = StringIds;
            this.contentTypes[key].callbacks.forEach(cb => cb(StringIds));
            this.contentTypes[key].callbacks = [];
          });
      }
    });
  }

  /**
   * Utility method: Get info for field
   */
  public getFieldInfo(
    {
      site,
      id,
      name,
      list
    }: { site?: string; list?: string } & (
      | { id: string; name?: string }
      | { id?: string; name: string }
    ),
    timeout: number = 10
  ) {
    if (site && !this.sites[site]) this.sites[site] = site;
    return Promise.all([
      list === undefined
        ? this.Web(site).then(web => web.contentTypes)
        : this.getListByTitle(list, site).then(library => library.contentTypes),
      this.getContentTypeIds({ site, list }),
      this.autoBatch(timeout, site)
    ])
      .then(([contentTypes, StringIds, getBatch]) =>
        Promise.all(
          StringIds.map(StringId =>
            contentTypes
              .getById(StringId)
              .fields.filter(
                id === undefined ? `InternalName eq '${name}'` : `Id eq '${id}'`
              )
              .inBatch(getBatch())
              .get()
          )
        )
      )
      .then(r => (r.find(test => test.length > 0) || [])[0] as IListField);
  }
}
