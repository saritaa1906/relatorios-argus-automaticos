import fs from "node:fs/promises";
import path from "node:path";
import { loadConfig } from "./config.mjs";
import { listMailings, fetchFourteenDays } from "./argus.mjs";
import { createDrive, fileExists, uploadReport } from "./drive.mjs";
import { buildReport, reportName } from "./report.mjs";

const config=loadConfig();
const today=new Date(); const todayIso=today.toLocaleDateString("en-CA",{timeZone:"America/Sao_Paulo"});
const mailings=await listMailings(config.argusToken);
const eligible=mailings.filter(m=>m.loteInicio?.slice(0,10)>=config.trackingDate).filter(m=>{const end=new Date(`${m.loteInicio.slice(0,10)}T12:00:00Z`);end.setUTCDate(end.getUTCDate()+13);return todayIso>=end.toISOString().slice(0,10);});
const drive=config.dryRun?null:createDrive(config); await fs.mkdir("outputs",{recursive:true}); await fs.mkdir("logs",{recursive:true});
const summary=[];
for(const mailing of eligible){
  const start=mailing.loteInicio.slice(0,10);const endDate=new Date(`${start}T12:00:00Z`);endDate.setUTCDate(endDate.getUTCDate()+13);const end=endDate.toISOString().slice(0,10);const name=reportName(mailing,start,end);
  if(!config.dryRun && await fileExists(drive,config.driveFolderId,name)){summary.push({lote:mailing.idLote,status:"já existe",name});continue;}
  const data=await fetchFourteenDays(config.argusToken,mailing);const output=path.resolve("outputs",name);await buildReport({mailing,rows:data.rows,start:data.start,end:data.end,outputPath:output});
  const uploaded=config.dryRun?null:await uploadReport(drive,config.driveFolderId,output,name);summary.push({lote:mailing.idLote,status:config.dryRun?"teste local":"enviado",name,registros:data.rows.length,driveId:uploaded?.id});
  await fs.rm(output,{force:true});
}
const result={executadoEm:new Date().toISOString(),dataSaoPaulo:todayIso,carteirasElegiveis:eligible.length,resultados:summary};await fs.writeFile("logs/ultima-execucao.json",JSON.stringify(result,null,2));console.log(JSON.stringify(result,null,2));

