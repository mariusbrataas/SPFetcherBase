import { IList, IField } from '@pnp/sp/presets/all';
import { ITerm, TaxonomyField, SPFetcherStructure } from './interfaces';
import { SPFetcherInitializer } from './SPFetcher.initializer';
import { SPFetcherBuildInterfaces } from './SPFetcher.buildInterfaces';

/**
 * SPFetcherBase
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

export class SPFetcherBase<
  T extends SPFetcherStructure
> extends SPFetcherBuildInterfaces<T> {
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

  public getFieldsInterfaceByListId(
    id: string | string[],
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return Promise.all(
      (id instanceof Array ? id : [id]).map(list_id =>
        this.getListById(list_id, site).then(list =>
          this.getFieldsInterface(list)
        )
      )
    ).then(r => r.join('\n\n'));
  }

  public getFieldsInterfaceByListTitle(
    titles: string | string[],
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return Promise.all(
      (titles instanceof Array ? titles : [titles]).map(title =>
        this.getListByTitle(title, site).then(list =>
          this.getFieldsInterface(list)
        )
      )
    ).then(r => r.join('\n\n'));
  }

  /**
   * Helper method: Get the taxonomy node closest to the requested path
   * @param termset_id
   * @param new_path
   */
  private getTaxonomyClosestParent(
    termset_id: string,
    new_path: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    const split_path = new_path.split(';');
    return this.getTermsetById(termset_id, site).then(
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
   * Helper method: Recursively add path to taxonomy termset
   * @param sspid
   * @param termset_id
   * @param parent_id
   * @param path
   */
  private addTaxonomyPath(
    sspid: string,
    termset_id: string,
    parent_id: string,
    ...path: string[]
  ): Promise<string> {
    return this.ready().then(() =>
      path.length
        ? this.post(
            '_vti_bin/taxonomyinternalservice.json/CreateTaxonomyItem',
            {
              body: JSON.stringify({
                sspId: sspid,
                lcid: 1033,
                parentType: termset_id === parent_id ? 3 : 4,
                webId: '00000000-0000-0000-0000-000000000000',
                listId: '00000000-0000-0000-0000-000000000000',
                parentId: parent_id,
                termsetId: termset_id,
                newName: path[0]
              })
            }
          )
            .then(r => r.json())
            .then(r =>
              this.addTaxonomyPath(
                sspid,
                termset_id,
                r.d.Content.Id,
                ...path.slice(1)
              )
            )
        : parent_id
    );
  }

  public buildTaxonomyPath(
    sspid: string,
    termset_id: string,
    new_path: string
  ) {
    const path = new_path;
    const split_path = path.split(';');
    return this.getTaxonomyClosestParent(
      termset_id,
      new_path
    ).then(([highscore, term]) =>
      this.addTaxonomyPath(
        sspid,
        termset_id,
        (term && term.Id) || termset_id,
        ...split_path.slice(highscore + 1)
      )
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
  public buildTaxonomyPathByFieldId(
    id: string,
    new_path: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getFieldById(id, site).then(field =>
      this.buildTaxonomyPathByField(field, new_path)
    );
  }

  /**
   * Utility method: Build taxonomy path at termset related to field by field title
   * @param title
   * @param new_path
   */
  public buildTaxonomyPathByFieldTitle(
    title: string,
    new_path: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getFieldByTitle(title, site).then(field =>
      this.buildTaxonomyPathByField(field, new_path)
    );
  }

  /**
   * Utility method: Build taxonomy path at termset related to field by field title
   * @param title
   * @param new_path
   */
  public buildTaxonomyPathByFieldInternalNameOrTitle(
    title: string,
    new_path: string,
    site?: keyof SPFetcherInitializer<T>['sites']
  ) {
    return this.getFieldByInternalNameOrTitle(title, site).then(field =>
      this.buildTaxonomyPathByField(field, new_path)
    );
  }
}

// const Fetcher = new SPFetcherBase();

// function getTypeOf(data, exclude = '', depth = 0) {
//   const type = typeof data;
//   return type === 'object'
//     ? data instanceof Array
//       ? `${getTypeOf(data[0], exclude, depth + 1)}[]`
//       : Object.keys(data || {}).length
//       ? `{${Object.keys(data || {})
//           .filter(test => test !== exclude)
//           .sort()
//           .reduce(
//             (prev, key) =>
//               key === 'odata.type'
//                 ? `${prev}\n  "${key}": "${data[key]}"`
//                 : `${prev}\n  ${
//                     typeof key === 'string'
//                       ? key.indexOf('.') === -1
//                         ? key
//                         : `"${key}"`
//                       : key
//                   }: ${getTypeOf(data[key], exclude, depth + 1)}`,
//             ''
//           )}\n}`
//       : 'any'
//     : type === 'undefined'
//     ? 'any'
//     : type;
// }

// function getInterfacesByList(lists) {
//   return Promise.all(lists.map(list => list.fields()))
//     .then(r => r.reduce((prev, current) => prev.concat(current), []))
//     .then(fields =>
//       fields.reduce((prev, field) => {
//         const type = field['odata.type'];
//         const prev_field = prev[type] || {};
//         return {
//           ...prev,
//           [type]: Object.keys(prev_field)
//             .concat(Object.keys(field))
//             .reduce(
//               (field_prev, prop) => ({
//                 ...field_prev,
//                 [prop]: prev_field[prop] || field[prop]
//               }),
//               {}
//             )
//         };
//       }, {})
//     )
//     .then(r =>
//       Object.keys(r)
//         .filter(key => key !== 'SP.Field')
//         .reduce(
//           (prev, key) => ({
//             ...prev,
//             [key]: Object.keys(r[key])
//               .sort()
//               .filter(test => !(test in r['SP.Field'] && test !== 'odata.type'))
//               .reduce(
//                 (prev_field, prop) => ({
//                   ...prev_field,
//                   [prop]: r[key][prop]
//                 }),
//                 {}
//               )
//           }),
//           { 'SP.Field': r['SP.Field'] }
//         )
//     )
//     .then(r => {
//       const base = r['SP.Field'];
//       return (
//         Object.keys(r)
//           .filter(key => key !== 'SP.Field')
//           .sort()
//           .reduce((prev, key) => {
//             const type = getTypeOf(r[key]);
//             return `${prev}\n\nexport interface ${key.replace(
//               /.*\./,
//               ''
//             )} extends Field ${type === 'any' ? '{}' : type}`;
//           }, `import { IFieldInfo } from "@pnp/sp/fields";\n\nexport interface Field extends IFieldInfo ${getTypeOf(base, 'odata.type')}`) +
//         '\n\nexport interface IListFields {' +
//         Object.keys(r)
//           .filter(key => key !== 'SP.Field')
//           .sort()
//           .map(key => key.replace(/.*\./, ''))
//           .reduce((prev, key) => `${prev}\n  ${key}: ${key};`, '') +
//         '\n}\n\nexport type IListField = IListFields[keyof IListFields];'
//       );
//     });
// }

// Promise.all(
//   ['SIOSAdmin', 'default'].map(site =>
//     Fetcher.Web(site)
//       .then(web => web.lists.get())
//       .then(lists =>
//         Promise.all(lists.map(list => Fetcher.getListByTitle(list.Title, site)))
//       )
//   )
// )
//   .then(lists => lists.flat())
//   .then(lists => getInterfacesByList(lists))
//   .then(console.log);

// Fetcher.web.lists
//   .get()
//   .then(r =>
//     Promise.all(
//       r
//         .map(({ Id }) => Id)
//         .concat('298a967c-2d33-4f01-9c7f-178354ec4720')
//         .map(Id => Fetcher.getListById(Id))
//     )
//   )
//   .then(lists => getInterfacesByList(...lists))
//   .then(console.log);
