const fs=require('fs');
const t=JSON.parse(fs.readFileSync('combined_extent.topojson','utf8'));
const [sx,sy]=t.transform.scale, [tx,ty]=t.transform.translate;
// decode arcs to absolute lon/lat
const arcs=t.arcs.map(arc=>{
  let x=0,y=0; return arc.map(([dx,dy])=>{ x+=dx; y+=dy; return [x*sx+tx, y*sy+ty]; });
});
function arcCoords(i){ return i>=0 ? arcs[i] : arcs[~i].slice().reverse(); }
function ringCoords(ringArcs){
  let pts=[];
  ringArcs.forEach((ai,k)=>{ const c=arcCoords(ai); pts=pts.concat(k>0?c.slice(1):c); });
  return pts;
}
function geomPolys(g){ // -> array of outer rings (lon/lat)
  const out=[];
  if(g.type==='Polygon'){ out.push(ringCoords(g.arcs[0])); }
  else if(g.type==='MultiPolygon'){ g.arcs.forEach(poly=> out.push(ringCoords(poly[0])) ); }
  return out;
}
function area(r){let a=0;for(let i=0;i<r.length-1;i++)a+=r[i][0]*r[i+1][1]-r[i+1][0]*r[i][1];return Math.abs(a/2);}
function dp(pts,eps){if(pts.length<4)return pts;function d(p,a,b){const dx=b[0]-a[0],dy=b[1]-a[1];const tt=((p[0]-a[0])*dx+(p[1]-a[1])*dy)/((dx*dx+dy*dy)||1);return Math.hypot(p[0]-(a[0]+tt*dx),p[1]-(a[1]+tt*dy));}
  function rec(s,e){let dm=0,idx=-1;for(let i=s+1;i<e;i++){const dd=d(pts[i],pts[s],pts[e]);if(dd>dm){dm=dd;idx=i;}}if(dm>eps)return rec(s,idx).slice(0,-1).concat(rec(idx,e));return[pts[s],pts[e]];}
  return rec(0,pts.length-1);}
function parseYear(ys){ const bc=ys.match(/([\d.]+)\s*B\.C\./); if(bc)return -parseInt(bc[1]); const ad=ys.match(/A\.D\.\s*([\d.]+)/); if(ad)return parseInt(ad[1]); return null; }

const ERAS={};
t.objects.CombinedExtentLayers_v6.geometries.forEach(g=>{
  const ys=g.properties.year_string;
  const yr=parseYear(ys); if(yr==null) return; // skip "All Periods"
  const rings=geomPolys(g).filter(r=>area(r)>0.03).map(r=>{
    let s=dp(r.map(p=>[+p[0].toFixed(4),+p[1].toFixed(4)]),0.04);
    if(s.length>3 && (s[0][0]!==s[s.length-1][0]||s[0][1]!==s[s.length-1][1])) s.push(s[0]);
    return s;
  }).filter(r=>r.length>3);
  ERAS[yr]=rings;
});
const keys=Object.keys(ERAS).map(Number).sort((a,b)=>a-b);
keys.forEach(y=>console.log('year '+String(y).padStart(5)+': '+ERAS[y].length+' rings, '+ERAS[y].reduce((s,r)=>s+r.length,0)+' pts'));
fs.writeFileSync('era_extents.json',JSON.stringify(ERAS));
console.log('periods:',keys.length,'| total KB:',(JSON.stringify(ERAS).length/1024|0));
