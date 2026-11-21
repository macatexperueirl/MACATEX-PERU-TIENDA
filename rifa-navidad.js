(function(){
  const API = 'https://script.google.com/macros/s/AKfycbzqc-adBTSJiL4hmtI1ZEerais9iQiG7cKMepm7rlxHhgDwG4c6QB1NmvHv5SqehLjbcg/exec';
  const WA = '51939843420';
  const YAPE = '930 523 654';
  const grid = document.getElementById('rifaGrid');
  if (!grid) return;
  const form = document.getElementById('rifaForm');
  const iframe = document.getElementById('rifaPostFrame');
  const numberInput = document.getElementById('rifaNumber');
  const selectedLabel = document.getElementById('rifaSelected');
  const statusText = document.getElementById('rifaStatus');
  let selected = '';
  let numbers = [];
  let submitting = false;

  function jsonp(){
    return new Promise((resolve,reject)=>{
      const cb='macatexRifa_'+Math.random().toString(36).slice(2,10);
      const s=document.createElement('script');
      const timer=setTimeout(()=>{cleanup();reject(new Error('Tiempo de espera agotado'));},12000);
      function cleanup(){clearTimeout(timer);delete window[cb];s.remove();}
      window[cb]=data=>{cleanup();resolve(data);};
      s.onerror=()=>{cleanup();reject(new Error('No se pudo consultar la disponibilidad'));};
      s.src=API+'?tipo=rifa&callback='+cb+'&t='+Date.now();
      document.head.appendChild(s);
    });
  }
  function render(){
    const counts={DISPONIBLE:0,RESERVADO:0,VENDIDO:0};
    grid.innerHTML='';
    numbers.forEach(item=>{
      const n=String(item.numero).padStart(3,'0');
      const state=String(item.estado||'DISPONIBLE').toUpperCase();
      counts[state]=(counts[state]||0)+1;
      const b=document.createElement('button');
      b.type='button'; b.className='rifa-number '+state.toLowerCase();
      b.textContent=n; b.disabled=state!=='DISPONIBLE';
      b.setAttribute('aria-label','Número '+n+' '+state.toLowerCase());
      if(n===selected)b.classList.add('selected');
      if(state==='DISPONIBLE')b.addEventListener('click',()=>selectNumber(n));
      grid.appendChild(b);
    });
    document.getElementById('rifaAvailable').textContent=counts.DISPONIBLE||0;
    document.getElementById('rifaReserved').textContent=counts.RESERVADO||0;
    document.getElementById('rifaSold').textContent=counts.VENDIDO||0;
    if(selected && !numbers.some(x=>String(x.numero).padStart(3,'0')===selected && String(x.estado).toUpperCase()==='DISPONIBLE')){
      selected=''; numberInput.value=''; selectedLabel.textContent='Ninguno'; statusText.textContent='El número seleccionado dejó de estar disponible. Elige otro.';
    }
  }
  function selectNumber(n){selected=n;numberInput.value=n;selectedLabel.textContent=n;statusText.textContent='Número '+n+' seleccionado. Completa tus datos para solicitar la reserva.';render();}
  async function refresh(){
    try{const data=await jsonp();if(!Array.isArray(data))throw new Error('Respuesta inválida');numbers=data;render();}
    catch(e){statusText.textContent=e.message+'. Intenta actualizar en unos segundos.';}
  }
  form.addEventListener('submit',e=>{
    e.preventDefault();
    if(submitting)return;
    if(!selected){statusText.textContent='Primero selecciona un número disponible.';return;}
    const current=numbers.find(x=>String(x.numero).padStart(3,'0')===selected);
    if(!current||String(current.estado).toUpperCase()!=='DISPONIBLE'){statusText.textContent='Ese número ya no está disponible. Actualiza el listado.';refresh();return;}
    const name=document.getElementById('rifaName').value.trim();
    const phone=document.getElementById('rifaPhone').value.trim();
    if(name.length<2||phone.replace(/\D/g,'').length<9){statusText.textContent='Revisa tu nombre y número de WhatsApp.';return;}
    submitting=true;
    const btn=document.getElementById('rifaSubmit');btn.disabled=true;btn.textContent='Enviando solicitud…';
    const payload={accion:'reservar_rifa',numero:selected,nombre:name,whatsapp:phone};
    const postForm=document.createElement('form');
    postForm.method='POST';postForm.action=API;postForm.target='rifaPostFrame';postForm.style.display='none';
    Object.entries(payload).forEach(([k,v])=>{const i=document.createElement('input');i.type='hidden';i.name=k;i.value=v;postForm.appendChild(i);});
    document.body.appendChild(postForm);postForm.submit();postForm.remove();
    const msg='Hola, MACATEXPERU. Solicito reservar el número '+selected+' de la Gran Rifa Navideña. Mi nombre: '+name+'. WhatsApp: '+phone+'. Realizaré el Yape de S/10 al '+YAPE+' y enviaré el comprobante para validación.';
    window.open('https://wa.me/'+WA+'?text='+encodeURIComponent(msg),'_blank','noopener');
    statusText.textContent='Solicitud enviada. Realiza el Yape de S/10 al '+YAPE+' y envía tu comprobante por WhatsApp. La reserva está sujeta a confirmación del sistema.';
    setTimeout(()=>{submitting=false;btn.disabled=false;btn.textContent='Solicitar reserva por S/10';refresh();},2500);
  });
  document.getElementById('rifaRefresh').addEventListener('click',refresh);
  refresh();setInterval(refresh,15000);
})();