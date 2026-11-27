const ALLOWED_ORIGIN = "https://1959414-5.github.io";
function json(data,status=200,origin=ALLOWED_ORIGIN){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store","Access-Control-Allow-Origin":origin,"Access-Control-Allow-Credentials":"true","Vary":"Origin"}})}
function cors(request){const origin=request.headers.get("Origin");if(origin!==ALLOWED_ORIGIN)return new Response("Forbidden origin",{status:403});return new Response(null,{status:204,headers:{"Access-Control-Allow-Origin":origin,"Access-Control-Allow-Credentials":"true","Access-Control-Allow-Methods":"GET,POST,OPTIONS","Access-Control-Allow-Headers":"Content-Type,Authorization","Access-Control-Max-Age":"86400","Vary":"Origin"}})}
function b64url(bytes){let binary="";for(const b of bytes)binary+=String.fromCharCode(b);return btoa(binary).replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"")}
function decodeB64url(s){s=s.replace(/-/g,"+").replace(/_/g,"/");s+="=".repeat((4-s.length%4)%4);return Uint8Array.from(atob(s),c=>c.charCodeAt(0))}
async function sign(data,secret){const key=await crypto.subtle.importKey("raw",new TextEncoder().encode(secret),{name:"HMAC",hash:"SHA-256"},false,["sign"]);return b64url(new Uint8Array(await crypto.subtle.sign("HMAC",key,new TextEncoder().encode(data))))}
async function makeToken(env,username,role){const payload=b64url(new TextEncoder().encode(JSON.stringify({sub:username,role:role,exp:Math.floor(Date.now()/1000)+28800})));const base="v1."+payload;return base+"."+await sign(base,env.SESSION_SECRET)}
async function verifyToken(token,env){if(!env.SESSION_SECRET)return null;const parts=String(token||"").split(".");if(parts.length!==3||parts[0]!=="v1")return null;const expected=await sign(parts[0]+"."+parts[1],env.SESSION_SECRET);if(expected!==parts[2])return null;try{const p=JSON.parse(new TextDecoder().decode(decodeB64url(parts[1])));if(!p.sub||!p.exp||p.exp<=Math.floor(Date.now()/1000))return null;return {sub:String(p.sub),role:["ADMIN","OPERACIONAL","FATURAMENTO"].includes(p.role)?p.role:"OPERACIONAL",exp:p.exp}}catch{return null}}
function bearer(request){return (request.headers.get("Authorization")||"").replace(/^Bearer\s+/i,"").trim()}
function contasConfiguradas(env){const contas=[];if(env.APP_ADMIN_USER&&env.APP_ADMIN_PASSWORD)contas.push({username:env.APP_ADMIN_USER,password:env.APP_ADMIN_PASSWORD,role:"ADMIN"});if(env.APP_FATURAMENTO_USER&&env.APP_FATURAMENTO_PASSWORD)contas.push({username:env.APP_FATURAMENTO_USER,password:env.APP_FATURAMENTO_PASSWORD,role:"FATURAMENTO"});if(env.APP_USER&&env.APP_PASSWORD)contas.push({username:env.APP_USER,password:env.APP_PASSWORD,role:"OPERACIONAL"});return contas}
function ambienteDaRequisicao(role,requested){return role==="ADMIN"&&String(requested||"").toUpperCase()==="TESTE"?"TESTE":"PRODUCAO"}
function acaoPermitida(role,acao,ambiente){
  if(role==="ADMIN"){if(ambiente==="TESTE")return new Set(["sugerirRota","obterConfigMotivoAtraso","salvarConfigMotivoAtraso","salvarConfigModulosFrota","salvarVeiculo","alterarSituacaoVeiculo","criarViagem","atualizarParadaViagem","atualizarStatusViagem","registrarOcorrenciaViagem","analisarXmlNFe","salvarNFeViagem","liberarViagem"]).has(acao);return true}
  if(ambiente!=="PRODUCAO")return false;
  if(role==="OPERACIONAL")return !new Set(["salvarConfigMotivoAtraso","salvarConfigModulosFrota","salvarVeiculo","alterarSituacaoVeiculo","criarViagem","analisarXmlNFe","salvarNFeViagem","removerNFeViagem","liberarViagem"]).has(acao);
  if(role==="FATURAMENTO")return new Set(["sugerirRota","criarViagem","analisarXmlNFe","salvarNFeViagem","removerNFeViagem","liberarViagem"]).has(acao);
  return false
}
export default {async fetch(request,env){
  const origin=request.headers.get("Origin");if(request.method==="OPTIONS")return cors(request);if(origin!==ALLOWED_ORIGIN)return new Response("Forbidden origin",{status:403});
  const url=new URL(request.url);
  if(url.pathname==="/health")return json({ok:true,service:"mapa-entregas-api",status:"online",authVersion:2});
  if(url.pathname==="/auth/login"&&request.method==="POST"){
    if(!env.SESSION_SECRET||!contasConfiguradas(env).length)return json({error:"Worker ainda não configurado. Secrets de autenticação pendentes."},503);
    let body;try{body=await request.json()}catch{return json({error:"JSON inválido"},400)}
    const conta=contasConfiguradas(env).find(c=>c.username===body.username&&c.password===body.password);
    if(!conta)return json({error:"Usuário ou senha inválidos"},401);
    return json({token:await makeToken(env,conta.username,conta.role),expiresIn:28800,user:conta.username,role:conta.role})
  }
  if(url.pathname==="/auth/validate"&&request.method==="GET"){const claims=await verifyToken(bearer(request),env);return claims?json({valid:true,user:claims.sub,role:claims.role}):json({valid:false},401)}
  if(url.pathname==="/public/api"&&request.method==="GET"){
    const formato=url.searchParams.get("formato")||"dados",permitidos=new Set(["dados","coresStatus","proximoDiaUtil","origemPerfinorte","localizacoesMotoristas"]);
    if(!permitidos.has(formato))return json({error:"Formato não permitido"},403);if(!env.GAS_URL||!env.GAS_SECRET)return json({error:"Integração Apps Script não configurada"},503);
    const target=new URL(env.GAS_URL);target.searchParams.set("formato",formato);target.searchParams.set("secret",env.GAS_SECRET);
    let upstream;try{upstream=await fetch(target.toString(),{redirect:"follow"})}catch(err){return json({error:"Falha ao consultar Apps Script",detail:String(err&&err.message||err)},502,origin)}
    return new Response(await upstream.text(),{status:upstream.status,headers:{"Content-Type":upstream.headers.get("Content-Type")||"application/json","Cache-Control":"no-store","Access-Control-Allow-Origin":origin,"Vary":"Origin"}})
  }
  if(url.pathname==="/api"||url.pathname.startsWith("/api/")){
    const claims=await verifyToken(bearer(request),env);if(!claims)return json({error:"Sessão inválida ou expirada"},401);if(!env.GAS_URL||!env.GAS_SECRET)return json({error:"Integração Apps Script não configurada"},503);
    const operation=url.pathname.slice(5),target=new URL(env.GAS_URL);
    if(operation)target.searchParams.set("formato",operation);
    if(request.method==="GET"){
      const ambiente=ambienteDaRequisicao(claims.role,url.searchParams.get("ambiente"));
      for(const [k,v] of url.searchParams){if(!["_authRole","_authUser","_ambiente","secret","ambiente"].includes(k))target.searchParams.set(k,v)}
      target.searchParams.set("secret",env.GAS_SECRET);target.searchParams.set("_authRole",claims.role);target.searchParams.set("_authUser",claims.sub);target.searchParams.set("_ambiente",ambiente);
      let upstream;try{upstream=await fetch(target.toString(),{redirect:"follow"})}catch(err){return json({error:"Falha ao consultar Apps Script",detail:String(err&&err.message||err)},502,origin)}
      return new Response(await upstream.text(),{status:upstream.status,headers:{"Content-Type":upstream.headers.get("Content-Type")||"application/json","Cache-Control":"no-store","Access-Control-Allow-Origin":origin,"Access-Control-Allow-Credentials":"true","Vary":"Origin"}})
    }
    if(request.method==="POST"){
      let body;try{body=await request.json()}catch{return json({error:"JSON inválido"},400)}
      const ambiente=ambienteDaRequisicao(claims.role,body._ambiente),acao=String(body.acao||"");
      if(!acaoPermitida(claims.role,acao,ambiente))return json({error:"Seu perfil não tem permissão para executar esta ação."},403);
      body._workerSecret=env.GAS_SECRET;body._authRole=claims.role;body._authUser=claims.sub;body._ambiente=ambiente;
      let upstream;try{upstream=await fetch(env.GAS_URL,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(body),redirect:"follow"})}catch(err){return json({error:"Falha ao consultar Apps Script",detail:String(err&&err.message||err)},502,origin)}
      return new Response(await upstream.text(),{status:upstream.status,headers:{"Content-Type":upstream.headers.get("Content-Type")||"application/json","Cache-Control":"no-store","Access-Control-Allow-Origin":origin,"Access-Control-Allow-Credentials":"true","Vary":"Origin"}})
    }
    return json({error:"Método não permitido"},405)
  }
  return json({error:"Endpoint não encontrado"},404)
}};