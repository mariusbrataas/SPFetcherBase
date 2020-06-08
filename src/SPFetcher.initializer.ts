import { sp, Web, IWeb } from '@pnp/sp/presets/all';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import {
  SPFetcherStructure,
  IFetcherBaseProperties,
  IFetcherPropertyTypes
} from './interfaces';

export class SPFetcherInitializer<T extends SPFetcherStructure> {
  // Properties
  public context: BaseComponentContext;
  protected web: IWeb;
  public status: 'not initialized' | 'initializing' | 'ready' | 'error';
  private queue: (() => void)[];

  // "Volatile" properties
  public urls: {
    [key in
      | T['urls']
      | IFetcherBaseProperties['urls']]: IFetcherPropertyTypes['urls'];
  };
  public lists: {
    [key in
      | T['lists']
      | IFetcherBaseProperties['lists']]: IFetcherPropertyTypes['lists'];
  };
  public termsets: {
    [key in
      | T['termsets']
      | IFetcherBaseProperties['termsets']]: IFetcherPropertyTypes['termsets'];
  };
  public sites: {
    [key in
      | T['sites']
      | IFetcherBaseProperties['sites']]: IFetcherPropertyTypes['sites'];
  };

  // Constructor
  constructor() {
    this.context = undefined;
    this.web = undefined;
    this.status = 'not initialized';
    this.queue = [];
    this.urls = {
      ...this.urls
    };
    this.lists = {
      ...this.lists
    };
    this.termsets = {
      ...this.termsets
    };
    this.sites = {
      ...this.sites
    };
  }

  /**
   * Initializer
   */
  public initialize(context: BaseComponentContext) {
    return new Promise(resolve => {
      if (this.status === 'initializing') {
        this.queue.push(resolve);
      } else {
        this.status = 'initializing';
        this.context = context;

        // Setup pnp
        sp.setup({ spfxContext: context });

        // Extract sites from context
        this.sites.default = this.context.pageContext.site.absoluteUrl;
        this.sites.current = `${this.sites.default}`;
        this.sites.base = `${this.sites.default}`.match(
          /(.*).sharepoint.com/
        )[0];

        // Extract urls from context
        this.urls.absolute = this.sites.default;
        this.urls.base = this.sites.base;

        // Get new web object
        this.getWeb()
          .then(web => (this.web = web))
          .then(() => resolve());
      }
    })
      .then(() => this.startupRoutines())
      .then(r => {
        this.status = 'ready';
        this.queue.forEach(callback => callback());
        this.queue = [];
        return r;
      })
      .catch((error: Error) => {
        this.status = 'error';
        throw error;
      });
  }

  /**
   * Get a new web object
   */
  public getWeb(site?: keyof SPFetcherInitializer<T>['sites']) {
    return this.ready().then(() =>
      site
        ? Web(this.sites[site])
        : Web(this.context.pageContext.site.absoluteUrl)
    );
  }

  /**
   * Get a new web object for the given site.
   * Only to be used by utility methods.
   */
  protected Web(site?: keyof SPFetcherInitializer<T>['sites']) {
    return this.ready().then(() => (site ? this.getWeb(site) : this.web));
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
   * Startup routines
   * This method will be called during fetcher initialization, before tasks in
   * the queue (i.e. tasks that called .ready() before initialization) are executed.
   */
  protected startupRoutines(): Promise<any> {
    return Promise.all([]);
  }

  /**
   * Set the absolute url to be used by the current fetcher instance.
   *
   * @param url
   */
  public setSite(site?: keyof SPFetcherInitializer<T>['sites']) {
    this.sites.current = this.sites[site] || this.sites.default;
    return this.getWeb(this.sites.current).then(web => (this.web = web));
  }
}
