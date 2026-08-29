let available=[],noStock=[],mode='available';
const g=document.getElementById('gallery'),s=document.getElementById('search'),c=document.getElementById('count'),
v=document.getElementById('viewer'),vi=document.getElementById('viewerImg'),vn=document.getElementById('viewerName');

function st(x){
  if(!x.stockLinked)return'Existencia: sin vincular';
  if(x.stock===null||x.stock==='')return'Existencia: sin dato';
  let n=Number(x.stock);
  if(!Number.isNaN(n)&&n<=0)return'AGOTADO';
  return'Existencia: '+x.stock
}
function cl(x){
  if(!x.stockLinked)return'stock unknown';
  let n=Number(x.stock);
  if(!Number.isNaN(n)&&n<=0)return'stock out';
  if(!Number.isNaN(n)&&n<=5)return'stock low';
  return'stock available'
}

function ensureTabs(){
  if(document.getElementById('catalogTabs'))return;
  const tabs=document.createElement('div');
  tabs.id='catalogTabs';
  tabs.innerHTML=`
    <button class="tab active" data-mode="available">Productos disponibles</button>
    <button class="tab" data-mode="nostock">Productos sin existencia</button>`;
  s.parentNode.insertBefore(tabs,s);
  tabs.addEventListener('click',e=>{
    const b=e.target.closest('.tab');
    if(!b)return;
    mode=b.dataset.mode;
    document.querySelectorAll('.tab').forEach(x=>x.classList.toggle('active',x===b));
    s.value='';
    s.placeholder=mode==='available'?'Buscar producto disponible...':'Buscar producto sin existencia...';
    render();
  });
}

async function load(){
  ensureTabs();
  let r=await fetch('catalog.json?t='+Date.now(),{cache:'no-store'});
  let d=await r.json();
  available=d.images||[];
  noStock=d.noStockImages||[];
  render()
}

function render(){
  const list=mode==='available'?available:noStock;
  let q=s.value.toLowerCase().trim();
  let f=list.filter(x=>(x.name+' '+(x.folder||'')).toLowerCase().includes(q));
  c.textContent=f.length+(f.length===1?' imagen':' imágenes');
  g.innerHTML='';

  for(const x of f){
    let e=document.createElement('article');
    e.className='card';

    if(mode==='available'){
      e.innerHTML=`<img loading="lazy" src="${x.src}"><div class="name">${x.name}</div><div class="${cl(x)}">${st(x)}</div>`;
      e.onclick=()=>{
        vi.src=x.src;
        vn.innerHTML=`${x.name}<div class="${cl(x)}">${st(x)}</div>`;
        v.showModal()
      };
    }else{
      e.innerHTML=`<img loading="lazy" src="${x.src}"><div class="name">${x.name}</div>`;
      e.onclick=()=>{
        vi.src=x.src;
        vn.textContent=x.name;
        v.showModal()
      };
    }
    g.appendChild(e)
  }

  if(!f.length){
    g.innerHTML=`<div class="empty">${mode==='available'?'No se encontraron productos.':'No hay productos sin existencia para mostrar.'}</div>`;
  }
}

s.oninput=render;
document.getElementById('refreshBtn').onclick=load;
document.getElementById('close').onclick=()=>v.close();
load();
setInterval(load,300000);
