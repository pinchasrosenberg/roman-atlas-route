const fs=require('fs');
const E=JSON.parse(fs.readFileSync('era_extents.json','utf8'));
const h=fs.readFileSync('index.html','utf8');
function grabArr(n){const s=h.indexOf('const '+n+' = [');const st=h.indexOf('[',s);let d=0,k=st;for(;k<h.length;k++){if(h[k]==='[')d++;else if(h[k]===']'){d--;if(d===0)break;}}return eval(h.slice(st,k+1));}
const PLACES=grabArr('PLACES'), ORGANIC=grabArr('PROV_POLY');
function pin(lon,lat,poly){let ins=false;for(let i=0,j=poly.length-1;i<poly.length;j=i++){const xi=poly[i][0],yi=poly[i][1],xj=poly[j][0],yj=poly[j][1];if(((yi>lat)!=(yj>lat))&&(lon<(xj-xi)*(lat-yi)/(yj-yi)+xi))ins=!ins;}return ins;}
function dseg(px,py,a,b){const dx=b[0]-a[0],dy=b[1]-a[1];const t=Math.max(0,Math.min(1,((px-a[0])*dx+(py-a[1])*dy)/((dx*dx+dy*dy)||1)));return Math.hypot(px-(a[0]+t*dx),py-(a[1]+t*dy));}
function dpoly(lon,lat,poly){let m=1/0;for(let i=0,j=poly.length-1;i<poly.length;j=i++)m=Math.min(m,dseg(lon,lat,poly[j],poly[i]));return m;}
function heOfPt(lon,lat){let pr=ORGANIC.find(x=>pin(lon,lat,x[1]));if(pr)return pr[0];let bd=1/0,best=null;ORGANIC.forEach(x=>{const d=dpoly(lon,lat,x[1]);if(d<bd){bd=d;best=x;}});return best[0];}
function centroid(r){let x=0,y=0,n=0;for(let i=0;i<r.length-1;i++){x+=r[i][0];y+=r[i][1];n++;}return[x/n,y/n];}
function labelRing(ring){const cnt={};PLACES.forEach(p=>{if(pin(p[3],p[4],ring)){const he=heOfPt(p[3],p[4]);cnt[he]=(cnt[he]||0)+1;}});let best=null,bc=0;for(const he in cnt)if(cnt[he]>bc){bc=cnt[he];best=he;}if(best)return best;const c=centroid(ring);return heOfPt(c[0],c[1]);}
function area(r){let a=0;for(let i=0;i<r.length-1;i++)a+=r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1];return Math.abs(a/2);}

const REGIONS=[]; // [year, provinceHe, ring]
Object.keys(E).map(Number).sort((a,b)=>a-b).forEach(yr=>{
  E[yr].forEach(ring=>{ if(area(ring)<0.05) return; REGIONS.push([yr, labelRing(ring), ring]); });
});
// stats
const byName={}; REGIONS.forEach(r=>byName[r[1]]=(byName[r[1]]||0)+1);
console.log('regions:',REGIONS.length);
console.log('provinces covered:',Object.keys(byName).length);
console.log('Egypt regions:',byName['מצרים']||0,'| Syria:',byName['סוריה']||0,'| Africa:',byName['אפריקה (פרוקונסולריס)']||0);
fs.writeFileSync('prov_regions.json',JSON.stringify(REGIONS));
console.log('KB:',(JSON.stringify(REGIONS).length/1024|0));
