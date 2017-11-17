**Konsole**

```
ng new ng-office-demo

npm start
```

**IE 11**

*Browser starten http://localhost:4200*

**Eclipse**

*Editor starten*

```
npm i --save @microsoft/office-js @types/office-js bootstrap ngx-bootstrap
```

**.angular-cli.json**

```
"assets": [
    ...
    { "glob": "**/*", "input": "../node_modules/@microsoft/office-js/dist", "output": "./assets/officejs" }
],

...

"styles": [
    "../node_modules/bootstrap/dist/css/bootstrap.min.css",
    ...
],
```

**src/index.html**

```
<head>
    ...
    <script src="./assets/officejs/office.js"  type="text/javascript"></script>
    ...
</head>
```

**manifest/ng-office-demo.xml**

```
<?xml version="1.0" encoding="UTF-8"?>

<OfficeApp
    xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
    xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
    xsi:type="TaskPaneApp">

    <Id>5f0629a1-a668-4cc6-b7e3-ea0b6b84ff3c</Id>
    <Version>1.0.0.0</Version>
    <ProviderName>Landeshauptstadt München</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <DisplayName DefaultValue="LHM Office Demo" />
    <Description DefaultValue="Demo für den Workshop."/>
    <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />
    <SupportUrl DefaultValue="https://wollmux.net " />

    <Hosts>
        <Host Name="Document" />
    </Hosts>

    <Requirements>
        <Sets DefaultMinVersion="1.1">
            <Set Name="DialogApi" />
        </Sets>
    </Requirements>
    <DefaultSettings>
        <SourceLocation DefaultValue="https://localhost:4200/" />
    </DefaultSettings>
    <Permissions>ReadWriteDocument</Permissions>
</OfficeApp>
```

**src/main.ts**

```
/// <reference path="../node_modules/@types/office-js/index.d.ts" />

...

if (window.hasOwnProperty('Office') && window.hasOwnProperty('Word')) {
    Office.initialize = function(reason) {
        // Schaltet die Telemetry von Office.js aus.
        OSF.Logger = null;
        platformBrowserDynamic().bootstrapModule(AppModule);
    };
}
```

**src/typings.d.ts**

```
...

declare var OSF;
```

**src/polyfills.ts**

*Polyfills für IE 11 aktivieren.*

**MS Office**

*Freigabe Ordner manifest*

*Eintragen Manifest-Pfad: Optionen/Trust Center/Trust Center Settings/Trusted Add-In Catalogs*

**package.json**

```
  "scripts": {
    ...
    "start": "ng serve --host 0.0.0.0 --ssl 1 --disable-host-check --sourcemaps=true",
    ...
  },
```

**Konsole**

```
npm start

ng generate component components/my-component
```

**IE 11**
*Browser starten https://localhost:4200*

*Zertifikat importieren*

**MS Office**

*MS Office starten*

*Add-In einfügen*

**src/app/app.module.ts**

```
...
import { RouterModule, Routes } from '@angular/router';
import { LocationStrategy, HashLocationStrategy } from '@angular/common';
...
const routes = [
    { path: 'my-component', component: MyComponentComponent }
];
...
@NgModule({
    ...
    imports: [
        ...
        RouterModule.forRoot(routes)
    ],
    ...
    providers: [
        ...
        { provide: LocationStrategy, useClass: HashLocationStrategy }
    ],
    ...
})
```

**src/app/app.component.html**

*alles löschen*

```
<div class="container">
    <div class="row">
        <div class="col-sm-1">
            <button class="btn btn-default" routerLink="/my-component">My Component</button>
        </div>
    </div>
    <div class="row">
        <div class="col-sm-12">
            <router-outlet></router-outlet>
        </div>
    </div>
</div>
```

**Konsole**

```
ng generate service services/office
```

**src/app/app.module.ts**

```
...
import { OfficeService } from './services/office.service';
...
@NgModule({
    ...
    providers: [
        ...
        OfficeService
    ],
    ...
})
```

**src/app/services/office.service.ts**

```
...
async insertText(text: string): Promise<void> {
    Word.run(context => {
        var doc = context.document;
        doc.body.insertText(text, 'Start');

        return context.sync();
    });
}
...
```

**src/app/components/my-component/my-component.component.ts**

```
...
import { OfficeService } from '../../services/office.service';
...
    constructor(private office: OfficeService) { }
...
    onClick() {
        this.office.insertText("Hello World!")
            .then(() => console.log("Finished!"));
    }
...
```

**src/app/components/my-component/my-component.component.html**

*alles löschen*

```
<div>
    <button class="btn btn-primary" (click)="onClick()">Insert Text</button>
</div>
```

**src/app/components/my-component/my-component.component.html**

```
<div>
    <input type="text" class="form-control" placeholder="Text" [value]="text" (input)="text = $event.target.value" />
    ...
</div>
```

**src/app/components/my-component/my-component.component.ts**

```
    ...
    private text: string = "Hello World!";
    ...
    onClick() {
        this.office.insertText(this.text)
        ...
    }
```
