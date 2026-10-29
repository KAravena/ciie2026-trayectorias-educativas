/* Offline HTML runtime. The .qmd renders with genuine Reveal.js when Quarto is installed. */
(()=>{'use strict';
 const slides=[...document.querySelectorAll('.slides>section.academic-slide')], total=slides.length;
 let index=0,speaker=null;
 function parseHash(){const m=location.hash.match(/^#\/?(\d+)/);return m?Math.min(total-1,Math.max(0,Number(m[1]))):0}
 function fit(){let scale=Math.min(innerWidth/1280,innerHeight/720);document.documentElement.style.setProperty('--scale',String(scale))}
 function status(){return {index:index,slide:slides[index],notes:slides[index]?.querySelector('.notes')?.textContent||''}}
 function noteBody(){if(!speaker||speaker.closed)return;const s=status();try{speaker.document.getElementById('speaker-title').textContent=`Diapositiva ${index+1} de ${total}`;speaker.document.getElementById('speaker-text').textContent=s.notes;speaker.document.getElementById('upcoming').textContent=index+1<total?slides[index+1].getAttribute('aria-label'):'Fin de la presentación';}catch(e){}}
 function go(n,hash=true){index=Math.max(0,Math.min(n,total-1));slides.forEach((s,i)=>{s.classList.toggle('present',i===index);s.setAttribute('aria-hidden',String(i!==index))});document.getElementById('count').textContent=`${index+1} / ${total}`;document.getElementById('prev').disabled=index===0;document.getElementById('next').disabled=index===total-1;if(hash)history.replaceState(null,'',`#/${index}`);noteBody()}
 function openSpeaker(){if(speaker&&!speaker.closed){speaker.focus();noteBody();return;}speaker=window.open('','ciie-speaker','width=1000,height=710');if(!speaker){alert('Permite ventanas emergentes para abrir las notas del expositor.');return;}speaker.document.write('<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Vista del expositor · CIIE 2026</title><style>body{font:20px/1.6 Arial,sans-serif;padding:28px;background:#f7f5fb;color:#2b2333}h1{color:#422168}pre{font:20px/1.65 Arial,sans-serif;white-space:pre-wrap;background:#fff;padding:22px;border-radius:9px}#upcoming{color:#7953a6}</style></head><body><h1 id="speaker-title"></h1><p>Notas originales del PowerPoint V29</p><pre id="speaker-text"></pre><h2>Próxima diapositiva</h2><p id="upcoming"></p><p>Controla la presentación con las flechas desde la ventana principal.</p></body></html>');speaker.document.close();noteBody()}
 function key(e){if(e.altKey||e.ctrlKey||e.metaKey)return;const t=e.target?.tagName;if(t==='INPUT'||t==='TEXTAREA')return;
  if(['ArrowRight','ArrowDown','PageDown',' ','Enter'].includes(e.key)){e.preventDefault();go(index+1)}
  else if(['ArrowLeft','ArrowUp','PageUp','Backspace'].includes(e.key)){e.preventDefault();go(index-1)}
  else if(e.key==='Home'){e.preventDefault();go(0)} else if(e.key==='End'){e.preventDefault();go(total-1)}
  else if(e.key.toLowerCase()==='s'){e.preventDefault();openSpeaker()}
  else if(e.key.toLowerCase()==='f'){e.preventDefault();if(!document.fullscreenElement)document.documentElement.requestFullscreen?.();else document.exitFullscreen?.()}
 }
 document.getElementById('prev').addEventListener('click',()=>go(index-1));document.getElementById('next').addEventListener('click',()=>go(index+1));
 addEventListener('resize',fit);addEventListener('keydown',key);addEventListener('hashchange',()=>go(parseHash(),false));
 fit();go(parseHash(),false);window.CIIERoyal={go,current:()=>index+1,total,notes:()=>status().notes};
})();
