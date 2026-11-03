import {requireActive,startMonitor,accountBar} from './auth-guard.js';
try{await requireActive();document.body.classList.remove('access-loading');accountBar(document.querySelector('.setup'));const script=document.createElement('script');script.src='setup.js';document.body.append(script);startMonitor();}catch{}
