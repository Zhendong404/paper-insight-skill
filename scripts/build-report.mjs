#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { parseArgs } from 'node:util';

const { values: opts } = parseArgs({options:{input:{type:'string'},output:{type:'string'},'runtime-modules':{type:'string'},'skill-dir':{type:'string'}}});
if (!opts.input || !opts.output || !opts.output.endsWith('.pptx')) throw new Error('Provide --input report.json and --output draft.pptx.');
const skillDir = opts['skill-dir'] ? path.resolve(opts['skill-dir']) : fileURLToPath(new URL('../', import.meta.url));
const manifest = JSON.parse(await fs.readFile(path.join(skillDir,'artifact-template.json'),'utf8'));
const reference = path.join(skillDir,manifest.reference);
const out = path.resolve(opts.output);
if (out === path.resolve(reference)) throw new Error('The reference PPTX is read-only.');
await fs.mkdir(path.dirname(out),{recursive:true});
const runtimeModules = opts['runtime-modules'];
const {FileBlob,PresentationFile} = runtimeModules
  ? await import(pathToFileURL(createRequire(path.join(path.resolve(runtimeModules),'_resolver.cjs')).resolve('@oai/artifact-tool')).href)
  : await import('@oai/artifact-tool');
const input = path.resolve(opts.input);
const report = JSON.parse(await fs.readFile(input,'utf8'));
function line(s,key) {
  if (typeof s !== 'string' || !s.trim() || /[\r\n]/.test(s)) throw new Error(`${key}: expected a non-empty single line.`);
  return s.trim();
}
function items(key,max) {
  const a=report[key];
  if (!Array.isArray(a)||a.length===0||a.length>max) throw new Error(`${key}: expected 1 to ${max} items.`);
  return a.map((s,i)=>line(s,`${key}[${i}]`));
}
const p=await PresentationFile.importPptx(await FileBlob.load(reference));
if (p.slides.items.length!==1) throw new Error('Expected the single-slide paper insight reference.');
const slide=p.slides.items[0];
const red='#C00000', black='#1D1D1A', dark='#151515', yahei='Microsoft YaHei', ui='Microsoft YaHei UI';
const run=(text,size,font,color,bold=false,italic=false)=>({run:text,style:{fontSize:`${size}pt`,typeface:font,color,bold,italic}});
const bullet='•', arrow='Ø'; // Wingdings Ø is the source template's arrow glyph.
const para=(r,marker,large=false)=>({runs:[r],paragraphStyle:{lineSpacingPercent:150000,...(marker?{bulletCharacter:marker,bulletTypeface:marker===arrow?'Wingdings':'Arial',bulletColorFollowsText:true,marginLeft:large?285750:171450,indent:large?-285750:-171450}:{})}});
const sub=(t,c=red,b=true)=>run(t,9,yahei,c,b);
function shape(id){
  const found=slide.shapes.items.find(x=>x.id===String(id));
  if (!found) throw new Error(`Missing source shape ${id}.`);
  return found;
}
function put(id,paragraphs,spaced=true){
  const s=shape(id);
  s.text.set(paragraphs);
  if (spaced) s.text.autoFit='none';
}
const title=line(report.title,'title');
put(2,[para(run(title,17,yahei,red,true))],false);
const insight=line(report.insight,'insight').replace(/^洞察\s*[|｜：:]\s*/, '');
const implication=line(report.implication,'implication').replace(/^(启示|启发)\s*[|｜：:]\s*/, '');
put(4,[para(run(`洞察 | ${insight}`,12,ui,black,true),bullet,true),para(run(`启示 | ${implication}`,12,ui,red,true),bullet,true)]);
put(13,[para(sub('①问题描述')),...items('problem',3).map((s,i)=>para(sub(s,i===0?red:black),arrow))]);
put(15,[para(sub('②'+line(report.method_title,'method_title').replace(/^②\s*/,''))),...items('method',5).map((s,i)=>para(sub(s,i===1?red:black,i===0),i===0?arrow:bullet))]);
const main=line(report.result_main,'result_main');
put(17,[para(sub('③技术收益')),para(sub(main,black),arrow)]);
if (report.result_secondary === null) slide.shapes.deleteById('22');
else put(22,[para(sub(line(report.result_secondary,'result_secondary')),arrow)]);
put(24,items('key_technologies',3).map(s=>para(run(s,12,ui,dark,true),bullet,true)));
put(26,[para(run(line(report.source,'source'),10,ui,dark,false,true))],false);
const figureSlots={context_figure:'14',method_figure:'16',evaluation_figure:'19',result_figure:'21'};
const figureNotes=[];
for (const [role,id] of Object.entries(figureSlots)) {
  if (!report.figures || !Object.hasOwn(report.figures,role)) throw new Error(`figures.${role}: provide a replacement or explicit null.`);
  const value=report.figures[role];
  const image=slide.images.items.find(x=>x.id===id);
  if (!image) throw new Error(`Missing source image ${id}.`);
  if (value===null){ image.delete(); continue; }
  const asset=path.resolve(path.dirname(input),line(value.path,`${role}.path`));
  const contentType={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg'}[path.extname(asset).toLowerCase()];
  if (!contentType) throw new Error(`${role}: use PNG or JPEG.`);
  const source=line(value.source,`${role}.source`);
  const alt=line(value.alt,`${role}.alt`);
  const frame={...image.frame};
  image.replace({blob:await fs.readFile(asset),contentType,alt,fit:'contain'});
  image.frame=frame;
  image.crop={left:0,top:0,right:0,bottom:0};
  image.fit='contain';
  figureNotes.push(`${alt}：${source}`);
}
if (typeof report.notes!=='string' || !report.notes.trim()) throw new Error('Provide notes with paper evidence and caveats.');
slide.speakerNotes.text=[report.notes.trim(),...figureNotes].join('\n');
const candidate=await PresentationFile.exportPptx(p);
await candidate.save(out);
const reopened=await PresentationFile.importPptx(await FileBlob.load(out));
const renderedSlide=reopened.slides.items[0];
if (reopened.slides.items.length!==1 || !renderedSlide.shapes.items.some(x=>x.text.toString()===title)) throw new Error('Export did not preserve the single slide and title.');
const stem=out.slice(0,-5);
const png=await reopened.export({slide:renderedSlide,format:'png',scale:1.5});
await fs.writeFile(stem+'.png',new Uint8Array(await png.arrayBuffer()));
const layout=await renderedSlide.export({format:'layout'});
await fs.writeFile(stem+'.layout.json',await layout.text());
console.log(out);
