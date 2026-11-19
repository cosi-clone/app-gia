(function(root) {
  'use strict';
  const indexed = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF '+
    '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF '+
    '800000 008000 000080 808000 800080 008080 C0C0C0 808080 '+
    '9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF '+
    '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF '+
    '00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 '+
    '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 '+
    '003366 339966 003300 333300 993300 993366 333399 333333').split(' ');
  const defaultTheme = ['FFFFFF','000000','EEECE1','1F497D','4F81BD','C0504D','9BBB59','8064A2','4BACC6','F79646','0000FF','800080'];
  function hex(value) {
    const clean = String(value || '').replace(/^#/, '');
    return /^(?:[a-f\d]{8}|[a-f\d]{6})$/i.test(clean) ? '#' + clean.slice(-6).toUpperCase() : '';
  }
  function tint(rgb, amount) {
    if (!rgb || !amount) return rgb;
    amount = Math.max(-1, Math.min(1, amount));
    const [r,g,b] = rgb.slice(1).match(/../g).map(x => parseInt(x,16)/255);
    const max = Math.max(r,g,b), min = Math.min(r,g,b), delta = max-min;
    let h=0, l=(max+min)/2, s=0;
    if (delta) {
      s=delta/(1-Math.abs(2*l-1));
      h=max===r ? ((g-b)/delta)%6 : max===g ? (b-r)/delta+2 : (r-g)/delta+4;
      h=(h*60+360)%360;
    }
    l=amount<0 ? l*(1+amount) : l*(1-amount)+amount;
    const c=(1-Math.abs(2*l-1))*s, x=c*(1-Math.abs((h/60)%2-1)), m=l-c/2;
    const values=h<60?[c,x,0]:h<120?[x,c,0]:h<180?[0,c,x]:h<240?[0,x,c]:h<300?[x,0,c]:[c,0,x];
    return '#'+values.map(v=>Math.round((v+m)*255).toString(16).padStart(2,'0')).join('').toUpperCase();
  }
  function resolve(element, theme=[], palette=[], fallback='') {
    if (!element) return fallback;
    let value='';
    if (element.hasAttribute('rgb')) value=hex(element.getAttribute('rgb'));
    else if (element.hasAttribute('theme')) {
      const i=+element.getAttribute('theme'); value=hex(theme[i]||defaultTheme[i]);
    } else if (element.hasAttribute('indexed')) {
      const i=+element.getAttribute('indexed');
      value=i===64?'#000000':i===65?'#FFFFFF':hex(palette[i]||indexed[i]);
    } else if (element.getAttribute('auto')==='1') value=fallback;
    return tint(value||fallback, +element.getAttribute('tint')||0);
  }
  root.XlsxColors={resolve,tint};
})(typeof window==='undefined'?globalThis:window);
