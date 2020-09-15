import { IList } from '@pnp/sp/presets/all';
import { SPFetcherStructure } from './interfaces';
import { SPFetcherUtils } from './SPFetcher.utils';

function fixTitle(title: string, capitalize?: boolean) {
  if (capitalize) title = `${title.slice(0, 1).toUpperCase()}${title.slice(1)}`;
  return title.replace(/site/g, 'Site');
}

function getTitle(title: string, capitalize?: boolean) {
  title = fixTitle(title, capitalize);
  return title.indexOf(' ') === -1 ? title : `"${title}"`;
}

function getEnumTitle(title: string) {
  title = fixTitle(title, true);
  return title.split(' ').join('');
}

function getInterfaceTitle(title: string, suffix?: string) {
  title = fixTitle(title, true);
  return `I${getEnumTitle(title)}${suffix || ''}`;
}

function SortByField<T extends { [key: string]: any }[], F extends keyof T[0]>(
  data: T,
  field: F
): T {
  return data
    ? data.sort((a, b) =>
        // @ts-ignore
        a[field] > b[field] ? 1 : a[field] === b[field] ? 0 : -1
      )
    : ([] as T);
}

function getDescription(descriptions: string[], prefix: string = '') {
  return [
    `/**`,
    ...descriptions.filter(test => test && test.length).map(msg => ` * ${msg}`),
    ` */`
  ]
    .map(msg => `${prefix}${msg}`)
    .join('\n');
}

function uniquify(items: string[]) {
  return Object.keys(
    items.reduce((prev, item) => ({ ...prev, [item]: true }), {})
  ).sort();
}

function createType(
  variant: 'enum' | 'interface',
  title: string,
  description: string,
  items: { title: string; type: string; description: string }[],
  suffix?: string
) {
  const name = `${(variant === 'enum'
    ? getEnumTitle
    : variant === 'interface'
    ? getInterfaceTitle
    : undefined)(title)}${suffix || ''}`;
  const items_typings = SortByField(items, 'title')
    .map(
      item =>
        `\n\n${getDescription(
          [...(item.description || item.title).split('\n')],
          '  '
        )}\n  ${getTitle(item.title, false)}${
          variant === 'enum' ? ' = ' : variant === 'interface' ? ': ' : ''
        }${item.type}${
          variant === 'enum' ? ',' : variant === 'interface' ? ';' : ''
        }`
    )
    .join('')
    .trim();
  return {
    imports: uniquify(items.map(item => item.type)),
    title,
    type_title: name,
    description,
    definition: `${getDescription([
      fixTitle(title, true),
      ...description.split('\n')
    ])}\nexport ${variant} ${name} {\n  ${items_typings}\n}`
  };
}

export class SPFetcherBuildInterfaces<
  T extends SPFetcherStructure
> extends SPFetcherUtils<T> {
  /**
   * Generate interface for the given list
   *
   * @param list
   */
  public getListTypings(list: IList) {
    return Promise.all([
      list.select('Title', 'Description').get(),
      list.fields.get().then(items =>
        items.map(item => ({
          title: item.InternalName,
          type: item['odata.type'].split('.').slice(-1)[0],
          description: item.Description
        }))
      )
    ]).then(([{ Title, Description }, items]) =>
      createType('interface', Title, Description, items, 'Item')
    );
  }

  /**
   * Generate interfaces for all selected lists on given site
   *
   * @param site_title
   * @param list_titles
   */
  public getSiteTypings(site_title: string, list_titles: string[]) {
    return Promise.all([
      this.Web(site_title).then(web =>
        web.select('Title', 'Description', 'Url').get()
      ),
      Promise.all(
        list_titles.map(list =>
          this.getListByTitle(list, site_title).then(list =>
            this.getListTypings(list)
          )
        )
      )
    ]).then(
      ([{ Title, Description: description, Url: url }, sub_definitions]) => ({
        url,
        sub_definitions,
        description,
        ...createType(
          'interface',
          site_title === 'default' ? site_title : Title,
          site_title === 'default' ? 'Default site' : description,
          sub_definitions.map(item => ({
            title: item.title,
            type: item.type_title,
            description: item.description
          }))
        )
      })
    );
  }

  /**
   * Generate interfaces for all given lists on all given sites
   *
   * @param sites
   */
  public getProjectTypings(sites: { [site: string]: string[] }) {
    return Promise.all(
      Object.keys(sites).map(site => this.getSiteTypings(site, sites[site]))
    )
      .then(items => ({
        imports: uniquify(
          items
            .map(item => item.sub_definitions.map(sub => sub.imports).flat())
            .flat()
        ).sort(),
        list_definitions: SortByField(
          items.map(item => item.sub_definitions).flat(),
          'type_title'
        ),
        sites_definitions: SortByField(items, 'type_title'),
        sites_enum: createType(
          'enum',
          'Sites',
          '',
          items.map(item => ({
            title: getEnumTitle(item.title),
            type: `"${item.title === 'default' ? 'default' : item.url}"`,
            description: item.description
          }))
        )
      }))
      .then(
        ({ imports, list_definitions, sites_definitions, sites_enum }) =>
          `import {\n${imports
            .map(msg => `\n  ${msg},`)
            .join('')}\n} from "spfetcherbase";\n\n${list_definitions
            .map(item => item.definition)
            .join('\n\n')}\n\n${sites_definitions
            .map(item => item.definition)
            .join('\n\n')}\n\n${sites_enum.definition}\n\n${
            createType(
              'interface',
              'Project',
              'Project data structure',
              sites_definitions.map(item => ({
                title: `[Sites.${getEnumTitle(item.title)}]`,
                type: item.type_title,
                description: item.description
              }))
            ).definition
          }`
      );
  }
}
