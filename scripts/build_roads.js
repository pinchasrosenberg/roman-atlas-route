const fs=require('fs');
const g=JSON.parse(fs.readFileSync(process.argv[2] || 'RomanRoadsWallsIntersect_v6.geojson','utf8'));
function parseYear(ys){const bc=ys.match(/([\d.]+)\s*B\.C\./);if(bc)return -parseInt(bc[1]);const ad=ys.match(/A\.D\.\s*([\d.]+)/);if(ad)return parseInt(ad[1]);return null;}
function dp(pts,eps){if(pts.length<3)return pts;function d(p,a,b){const dx=b[0]-a[0],dy=b[1]-a[1];const t=((p[0]-a[0])*dx+(p[1]-a[1])*dy)/((dx*dx+dy*dy)||1);return Math.hypot(p[0]-(a[0]+t*dx),p[1]-(a[1]+t*dy));}
  function rec(s,e){let dm=0,idx=-1;for(let i=s+1;i<e;i++){const dd=d(pts[i],pts[s],pts[e]);if(dd>dm){dm=dd;idx=i;}}if(dm>eps)return rec(s,idx).slice(0,-1).concat(rec(idx,e));return[pts[s],pts[e]];}
  return rec(0,pts.length-1);}
// keep major roads (M) + walls (F); drop minor (n). group by earliest year they appear.
function llen(pts){let d=0;for(let i=0;i<pts.length-1;i++)d+=Math.hypot(pts[i+1][0]-pts[i][0],pts[i+1][1]-pts[i][1]);return d;}
const OUT=[]; let kept=0,pts=0;
g.features.forEach(f=>{
  const cl=f.properties.CLASS; if(cl!=='M') return;   // major roads only
  const yr=parseYear(f.properties.year_string); if(yr==null) return;
  f.geometry.coordinates.forEach(line=>{
    if(llen(line)<0.18) return;                         // drop very short segments
    let s=dp(line.map(p=>[+p[0].toFixed(2),+p[1].toFixed(2)]),0.05);
    if(s.length<2) return;
    OUT.push([yr, s]);  // [year, polyline]
    kept++; pts+=s.length;
  });
});
console.log('kept lines:',kept,'pts:',pts);
fs.writeFileSync('roads_filtered.json',JSON.stringify(OUT));
console.log('KB:',(JSON.stringify(OUT).length/1024|0));
