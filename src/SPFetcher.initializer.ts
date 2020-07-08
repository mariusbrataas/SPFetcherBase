import { sp, Web, IWeb, SPBatch } from '@pnp/sp/presets/all';
import { BaseComponentContext } from '@microsoft/sp-component-base';
import {
  SPFetcherStructure,
  IFetcherBaseProperties,
  IFetcherPropertyTypes
} from './interfaces';

export class SPFetcherInitializer<T extends SPFetcherStructure> {
  // Properties
  public context: BaseComponentContext;
  public status: 'not initialized' | 'initializing' | 'ready' | 'error';
  private queue: (() => void)[];
  private batches: {
    [key in T['sites'] | IFetcherBaseProperties['sites']]?: {
      batch: SPBatch;
      count: number;
    };
  };
  private webs: {
    [key in T['sites'] | IFetcherBaseProperties['sites']]: IWeb;
  };

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
    this.status = 'not initialized';
    this.queue = [];
    this.batches = {};
    this.webs = {
      ...this.webs
    };
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
        this.Web().then(() => resolve());
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
   * Get the sp object
   */
  public SP() {
    return this.ready().then(() => sp);
  }

  /**
   * Get a new web object for the given site.
   * Only to be used by utility methods.
   */
  public Web(site: keyof SPFetcherInitializer<T>['sites'] = 'default') {
    return this.ready().then(
      () => (this.webs[site] = this.webs[site] || Web(this.sites[site]))
    );
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
   * Utility method: Initialize empty batch object
   */
  public createBatch(site: keyof SPFetcherInitializer<T>['sites'] = 'default') {
    return this.Web(site).then(web => web.createBatch());
  }

  /**
   * Utility method: Returns a batch object that will execute automatically
   *
   * @param timeout - Maximum delay before executing batch
   */
  public autoBatch(
    timeout: number = 50,
    site: keyof SPFetcherInitializer<T>['sites'] = 'default'
  ) {
    const key = `[${timeout}ms] - ${site}`;
    return this.Web(site).then(web => {
      return () => {
        if (this.batches[key] && this.batches[key].count >= 50)
          this.batches[key] = undefined;
        if (!this.batches[key]) {
          const batch = web.createBatch();
          this.batches[key] = {
            batch,
            count: 0
          };
          setTimeout(() => {
            this.batches[key] = undefined;
            setTimeout(() => {
              batch.execute();
            }, 5);
          }, timeout);
        }
        this.batches[key].count += 1;
        return this.batches[key].batch;
      };
    });
  }
}
