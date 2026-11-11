import {inspect} from './source.mjs';
import {readFixture} from './files.mjs';
const name=process.argv[2]??'disclosure-1900.xlsx';
console.log(JSON.stringify(inspect(await readFixture(name)),null,2));
