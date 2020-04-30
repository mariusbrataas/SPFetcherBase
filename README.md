# SPFetcherBase

A container for all requests in a Sharepoint webpart

## Usage

- Start by extending this class and create your own constructor if necessary.
- Put startup-routines in the method "startupRoutines". This will be called automatically during initialization.
- Start your own methods with return this.ready().then(() => { ... }). This ensures execution once ready.
- While loading your webpart, call Fetcher.initialize to set it up

## Example

### Step 1: Create a new fetcher

```ts
class MyFetcherClass extends SPFetcherBase {
  // Properties
  protected urls: {
    absolute: string;
    logic: string;
  };

  // Constructor
  constructor() {
    super();
    this.urls = {
      absolute: undefined,
      logic: undefined
    };
  }

  // Startup routines
  protected startupRoutines(): Promise<any> {
    return Promise.all([
      this.web
        .getStorageEntity('SettingsSite')
        .then((r: any) =>
          new Web(r.Value).lists
            .getByTitle('Settings')
            .items.filter("Title eq 'CopyFileLogicAppUrl'")
            .get()
        )
        .then(r => (this.urls.logic = r.value))
    ]);
  }

  // Example method: get
  public getFiles() {
    return this.getDefaultLibrary().then(library =>
      library.items.filter(`startswith(ContentTypeId,'0x0101')`).get()
    );
  }
}

export const MyFetcher = new MyFetcherClass();
```

### Step 2: Initialize the fetcher from your webpart

```ts
import { MyFetcher } from './MyFetcher';

class MyWebpart {
  constructor({ context }) {
    MyFetcher.initialize(context);
  }

  /* ... */
}
```
