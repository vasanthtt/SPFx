## SPFx using react-adal package
SPFx webpart sample accessing secured custom and graph API using react-adal (https://www.npmjs.com/package/react-adal)

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp serve
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* sharepoint/* - all resources which should be uploaded to SharePoint Apps under tenant or sitecollection.

### Deploy the code

```bash
gulp clean
gulp build
gulp bundle --ship
gulp package-solution --ship
```
