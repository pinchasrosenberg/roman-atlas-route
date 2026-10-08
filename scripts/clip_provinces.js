const fs=require('fs');
const E=JSON.parse(fs.readFileSync('era_extents.json','utf8'));
const h=fs.readFileSync('index.html','utf8');
function grabArr(n){const s=h.indexOf('const '+n+' = [');const st=h.indexOf('[',s);let d=0,k=st;for(;k<h.length;k++){if(h[k]==='[')d++;else if(h[k]===']'){d--;if(d===0)break;}}return eval(h.slice(st,k+1));}
const ORGANIC=grabArr('PROV_POLY');

// ---------- Greiner-Hormann polygon clipping (intersection) ----------
function ptInPoly(p,poly){let ins=false;for(let i=0,j=poly.length-1;i<poly.length;j=i++){const xi=poly[i].x,yi=poly[i].y,xj=poly[j].x,yj=poly[j].y;if(((yi>p.y)!=(yj>p.y))&&(p.x<(xj-xi)*(p.y-yi)/(yj-yi)+xi))ins=!ins;}return ins;}
function toVerts(ring){return ring.map(c=>({x:c[0],y:c[1]}));}
function clip(subjRing, clipRing){
  // returns array of result rings (subj ∩ clip)
  const S=toVerts(subjRing.slice(0,-1)), C=toVerts(clipRing.slice(0,-1));
  // build linked lists
  function build(arr){const n=arr.length;const v=arr.map(p=>({x:p.x,y:p.y,inter:false,alpha:0,entry:false,neighbor:null,visited:false}));
    for(let i=0;i<n;i++){v[i].next=v[(i+1)%n];v[i].prev=v[(i-1+n)%n];}return v;}
  let subj=build(S), clp=build(C);
  function insertBetween(a,b,node){ // insert node into ring between a..b by alpha order along a->... to b
    let cur=a; while(cur!==b && cur.next.inter && cur.next.alpha<node.alpha) cur=cur.next; node.next=cur.next;node.prev=cur;cur.next.prev=node;cur.next=node; }
  let anyInter=false;
  // find intersections
  const subjStart=subj.slice(); // original vertices
  for(let i=0;i<subjStart.length;i++){
    const a=subjStart[i], b=subjStart[(i+1)%subjStart.length];
  }
  // iterate original edges via arrays
  function edges(vs){const out=[];let start=vs[0];/* use original array order */return out;}
  // Simpler: iterate original arrays S,C for edges, but insert into linked list
  const subjArr=subj, clpArr=clp;
  for(let i=0;i<subjArr.length;i++){
    const a1=subjArr[i], a2=subjArr[(i+1)%subjArr.length];
    for(let j=0;j<clpArr.length;j++){
      const b1=clpArr[j], b2=clpArr[(j+1)%clpArr.length];
      const d=(b2.y-b1.y)*(a2.x-a1.x)-(b2.x-b1.x)*(a2.y-a1.y);
      if(Math.abs(d)<1e-12) continue;
      const ua=((b2.x-b1.x)*(a1.y-b1.y)-(b2.y-b1.y)*(a1.x-b1.x))/d;
      const ub=((a2.x-a1.x)*(a1.y-b1.y)-(a2.y-a1.y)*(a1.x-b1.x))/d;
      if(ua<=1e-9||ua>=1-1e-9||ub<=1e-9||ub>=1-1e-9) continue;
      const ix=a1.x+ua*(a2.x-a1.x), iy=a1.y+ua*(a2.y-a1.y);
      const sN={x:ix,y:iy,inter:true,alpha:ua,entry:false,neighbor:null,visited:false};
      const cN={x:ix,y:iy,inter:true,alpha:ub,entry:false,neighbor:null,visited:false};
      sN.neighbor=cN; cN.neighbor=sN;
      insertBetween(a1,a2,sN); insertBetween(b1,b2,cN); anyInter=true;
    }
  }
  if(!anyInter){
    // no crossing: either subj fully inside clip, or disjoint
    if(ptInPoly(S[0],C)) return [subjRing];
    return [];
  }
  // mark entry/exit for subject relative to clip
  function markEntry(ringHead, other){
    let status = ptInPoly(ringHead, otherVerts(other)) ? false : true; // if start inside other -> first intersection is 'exit'(false) else 'entry'(true)
    let cur=ringHead;
    do{ if(cur.inter){ cur.entry=status; status=!status; } cur=cur.next; }while(cur!==ringHead);
  }
  function otherVerts(vs){ // reconstruct polygon vertex list (non-inter original) from linked list head
    const arr=[]; let cur=vs; do{arr.push({x:cur.x,y:cur.y}); cur=cur.next;}while(cur!==vs); return arr; }
  const subjHead=subjArr[0], clpHead=clpArr[0];
  markEntry(subjHead, clpHead);
  markEntry(clpHead, subjHead);
  // trace
  const result=[];
  let guard=0;
  // collect all intersection nodes on subject
  const interNodes=[]; { let cur=subjHead; do{ if(cur.inter&&!cur.visited) interNodes.push(cur); cur=cur.next;}while(cur!==subjHead); }
  interNodes.forEach(startN=>{
    if(startN.visited) return;
    const poly=[]; let cur=startN;
    do{
      cur.visited=true; if(cur.neighbor) cur.neighbor.visited=true;
      poly.push([cur.x,cur.y]);
      if(cur.entry){ do{ cur=cur.next; poly.push([cur.x,cur.y]); }while(!cur.inter); }
      else { do{ cur=cur.prev; poly.push([cur.x,cur.y]); }while(!cur.inter); }
      cur=cur.neighbor;
      if(++guard>100000) break;
    }while(cur!==startN && !cur.visited);
    if(poly.length>3){ poly.push(poly[0]); result.push(poly); }
  });
  return result;
}
function area(r){let a=0;for(let i=0;i<r.length-1;i++)a+=r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1];return Math.abs(a/2);}
function bbox(r){let b=[1/0,1/0,-1/0,-1/0];r.forEach(p=>{const x=p[0]!==undefined?p[0]:p.x,y=p[1]!==undefined?p[1]:p.y;b[0]=Math.min(b[0],x);b[1]=Math.min(b[1],y);b[2]=Math.max(b[2],x);b[3]=Math.max(b[3],y);});return b;}
function bboxOverlap(a,b){return !(a[2]<b[0]||b[2]<a[0]||a[3]<b[1]||b[3]<a[1]);}

// clip each organic province against each era ring; keep fragments tagged with ring year
const OUT=[]; // [year, name, ring]
const eras=Object.keys(E).map(Number).sort((a,b)=>a-b);
ORGANIC.forEach(([name,poly])=>{
  const pb=bbox(poly);
  eras.forEach(yr=>{
    E[yr].forEach(ring=>{
      const rb=bbox(ring); if(!bboxOverlap(pb,rb)) return;
      let frags;
      try{ frags=clip(poly, ring); }catch(e){ frags=[]; }
      frags.forEach(f=>{ if(area(f)>0.02) OUT.push([yr,name,f.map(p=>[+p[0].toFixed(3),+p[1].toFixed(3)])]); });
    });
  });
});
const byName={}; OUT.forEach(o=>byName[o[1]]=(byName[o[1]]||0)+1);
console.log('clipped fragments:',OUT.length,'| provinces:',Object.keys(byName).length,'| total pts:',OUT.reduce((s,o)=>s+o[2].length,0));
console.log('Egypt frags:',byName['מצרים'],'Syria:',byName['סוריה'],'Africa:',byName['אפריקה (פרוקונסולריס)'],'Maur:',byName['מאוריטניה']);
fs.writeFileSync('prov_clipped.json',JSON.stringify(OUT));
console.log('KB:',(JSON.stringify(OUT).length/1024|0));
