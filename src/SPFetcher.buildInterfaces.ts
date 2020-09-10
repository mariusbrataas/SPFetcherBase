import { IList } from '@pnp/sp/presets/all';
import { SPFetcherStructure } from './interfaces';
import { SPFetcherUtils } from './SPFetcher.utils';

function getDescription(descriptions: string[], prefix: string = '') {
  return [
    `/** ${descriptions[0]}`,
    ...descriptions
      .slice(1)
      .filter(test => test && test.length)
      .map(msg => ` * ${msg}`),
    ` */`
  ]
    .map(msg => `${prefix}${msg}`)
    .join('\n');
}

interface IDefinition {
  title: string;
  interface_title: string;
  definition: string;
  imports: string[];
  description: string[];
}

export class SPFetcherBuildInterfaces<
  T extends SPFetcherStructure
> extends SPFetcherUtils<T> {
  /**
   * Generate interface for the given list
   *
   * @param list
   */
  public getListTypings(list: IList): Promise<IDefinition> {
    return list
      .select('Title')
      .get()
      .then(({ Title }) => Title.split(' ').join(''))
      .then(title => ({ title, interface_title: `I${title}Item` }))
      .then(({ title, interface_title }) =>
        list
          .fields()
          .then(fields =>
            fields.reduce(
              (prev, field) => ({
                ...prev,
                [field.InternalName]: {
                  type: field['odata.type'].split('.').slice(-1)[0],
                  description: getDescription(
                    [
                      `${field.InternalName} - [${field.Title}]`,
                      ...(field.Description || '').split('\n')
                    ],
                    '  '
                  )
                }
              }),
              {} as {
                [key: string]: { type: string; description: string };
              }
            )
          )
          .then(r =>
            Promise.all([
              {
                definition: Object.keys(r)
                  .map(
                    key =>
                      `\n\n${r[key].description}\n  ${key}: ${r[key].type};`
                  )
                  .sort()
                  .join(''),
                imports: Object.keys(
                  Object.keys(r).reduce(
                    (prev, key) => ({ ...prev, [r[key].type]: true }),
                    {}
                  )
                )
              },
              list
                .select('Title', 'Description')
                .get()
                .then(({ Title, Description }) => [Title, Description])
            ])
          )
          .then(([{ definition, imports }, description]) => ({
            title,
            interface_title,
            definition: `${getDescription(
              description
            )}\nexport interface ${interface_title} {${definition}\n}`,
            imports,
            description
          }))
      );
  }

  /**
   * Generate interfaces for all selected lists on given site
   *
   * @param site
   * @param lists
   */
  public getSiteTypings(site: string, lists: string[]): Promise<IDefinition> {
    return this.Web(site as any)
      .then(web => web.select('Title').get())
      .then(({ Title }) => Title.split(' ').join(''))
      .then(title => ({
        title,
        interface_title: `I${title}${
          title.toLowerCase().indexOf('site') === -1 ? 'Site' : ''
        }`.replace(/site/g, 'Site')
      }))
      .then(({ title, interface_title }) =>
        Promise.all(
          lists.map(list_title =>
            this.getListByTitle(list_title, site as any).then(list =>
              this.getListTypings(list)
            )
          )
        ).then(r =>
          Promise.all([
            Object.keys(
              r
                .map(typings => typings.imports)
                .flat()
                .reduce((prev, key) => ({ ...prev, [key]: true }), {})
            ).sort(),
            r
              .map(
                ({ title: tmp_title, interface_title, description }) =>
                  `\n\n${getDescription(
                    description,
                    '  '
                  )}\n  ${tmp_title}: ${interface_title};`
              )
              .sort()
              .join(''),
            this.Web(site as any)
              .then(web => web.select('Title', 'Description').get())
              .then(({ Title, Description }) => [Title, Description])
          ]).then(([imports, definition, description]) => ({
            title,
            interface_title,
            definition: `${r
              .map(typings => typings.definition)
              .sort()
              .join('\n\n')}\n\n${getDescription(
              description
            )}\nexport interface ${interface_title} {${definition}\n}`,
            imports,
            description
          }))
        )
      );
  }

  /**
   * Generate interfaces for all given lists on all given sites
   *
   * @param sites
   * @param include_imports
   */
  public getProjectTypings(
    sites: { [site: string]: string[] },
    include_imports?: boolean
  ): Promise<IDefinition> {
    const title = 'Project';
    const interface_title = `I${title}`;
    return Promise.all(
      Object.keys(sites).map(site_title =>
        this.getSiteTypings(site_title || 'default', sites[site_title])
      )
    ).then(r =>
      Promise.all([
        Object.keys(
          r
            .map(typings => typings.imports)
            .flat()
            .reduce((prev, key) => ({ ...prev, [key]: true }), {})
        ).sort(),
        r
          .map(
            ({ title: tmp_title, interface_title, description }) =>
              `\n\n${getDescription(
                description,
                '  '
              )}\n  ${tmp_title}: ${interface_title};`
          )
          .sort()
          .join(''),
        [`Project typings`]
      ])
        .then(([imports, definition, description]) => ({
          definition: `${r
            .map(typings => typings.definition)
            .sort()
            .join('\n\n')}\n\n${getDescription(
            description
          )}\nexport interface ${interface_title} {${definition}\n}`,
          imports,
          description
        }))
        .then(({ definition, imports, description }) => ({
          title,
          interface_title,
          definition: include_imports
            ? `${getDescription([
                'Import definitions from spfetcherbase'
              ])}\nimport {${imports.map(
                imp => `\n  ${imp}`
              )}\n} from "spfetcherbase";\n\n${definition}`
            : definition,
          imports,
          description
        }))
    );
  }
}
