import { sp } from "@pnp/sp-commonjs";
import { SPFetchClient } from '@pnp/nodejs-commonjs';

console.log('test');

sp.setup({
  sp: {
    fetchClientFactory: () => {
        return new SPFetchClient("{ site url }", "{ client id }", "{ client secret }");
    },
  },
});

console.log('test 2');
