import http from 'node:http';
const PORT=process.env.PORT||8787;
const server=http.createServer((req,res)=>{
  res.setHeader('Access-Control-Allow-Origin','*');
  res.setHeader('Access-Control-Allow-Headers','Content-Type');
  if(req.method==='OPTIONS'){res.writeHead(204);return res.end();}
  if(req.method!=='POST'||req.url!=='/verify'){res.writeHead(404);return res.end('Not found');}
  let body=''; req.on('data',c=>{body+=c;if(body.length>1_000_000)req.destroy();});
  req.on('end',()=>{
    try{
      const input=JSON.parse(body||'{}');
      const refs=Array.isArray(input.references)?input.references:[];
      const results=refs.map(r=>({id:r.id,status:'не проверено',raw:r.raw,note:'Подключите официальный источник/провайдер правовой верификации. ВНД не сохраняется.'}));
      res.writeHead(200,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify({results}));
    }catch(e){res.writeHead(400,{'Content-Type':'application/json'});res.end(JSON.stringify({error:'bad request'}));}
  });
});
server.listen(PORT,()=>console.log(`Legal Verification demo: http://localhost:${PORT}/verify`));
