import {readdir} from 'node:fs/promises';
import {spawnSync} from 'node:child_process';
async function check(directory) {
 for(const file of await readdir(directory,{withFileTypes:true})) {
  const name=`${directory}/${file.name}`;
  if(file.isDirectory()) await check(name);
  else if(name.endsWith('.mjs')) {const p=spawnSync(process.execPath,['--check',name],{stdio:'inherit'});if(p.status!==0)process.exit(p.status??1);}
 }
}
for(const folder of ['src','tests','tools']) await check(folder);
console.log('语法检查通过');
