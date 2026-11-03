import { build } from 'esbuild';
await build({entryPoints:['client/firebase.js'],outfile:'public/firebase-client.js',bundle:true,format:'esm',platform:'browser',target:'es2020',minify:true,legalComments:'eof'});
console.log('Đã build Firebase client.');
