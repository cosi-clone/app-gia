import {requireActive,startMonitor} from './auth-guard.js';
try{await requireActive();document.body.classList.remove('access-loading');const script=document.createElement('script');script.src='reader.js';document.body.append(script);startMonitor();}catch{}
