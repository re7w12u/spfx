# sp-pnp-webpart

## Summary

just a dummy example on how to fetch data from a SP list using pnp.

the trick here, is the pnpjsConfig.ts class.
It's first initialized in the onInit method of the GetListItemsPnpWebPart.ts component with the context. 
Then it's directly available for any other component without the need to use props or whatever. 

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.19.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
