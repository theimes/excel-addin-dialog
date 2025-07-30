# Readme

## Develop

Start the development server in a terminal

```{shell}
npm run dev-server
```

Set link to document. Get the link from the `Share` -> `Copy Link` ribbon in Excel web app.

```{shell}
export DOCLINK="https://my.sharepoint.com<YOUR_DOCUMENT_LINK>"
```


Activate the Add-In:

```{shell}
npm run start -- web --document $DOCLINK
```

Development certificates can be found at

```{shell}
/home/<USER>/.office-addin-dev-certs/ca.crt
/home/<USER>/.office-addin-dev-certs/localhost.crt
/home/<USER>/.office-addin-dev-certs/localhost.key
```
