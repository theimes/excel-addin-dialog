# Readme

## Develop

Start the development server in a terminal

```{shell}
npm run dev-server
```

Set link to document. Get the link from the `Share` -> `Copy Link` ribbon in Excel web app.

```{shell}
export DOCLINK="https://bergzeit-my.sharepoint.com/:x:/r/personal/thorsten_heimes_bergzeit_de/Documents/Demo.xlsx?d=wc503136d5d4b4bfcb6ce9ec1301bd8c1&csf=1&web=1&e=8d6k78"
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
