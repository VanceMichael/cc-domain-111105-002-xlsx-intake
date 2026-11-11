export function listZipEntries(input) {
  const b = Buffer.from(input); let end = -1;
  for (let i=b.length-22;i>=Math.max(0,b.length-65557);i--) {
    if (b.readUInt32LE(i)===0x06054b50 && i+22+b.readUInt16LE(i+20)===b.length) {end=i;break;}
  }
  if (end<0) throw new Error('缺少 ZIP 结束目录');
  if (b.readUInt16LE(end+4)!==0 || b.readUInt16LE(end+6)!==0) throw new Error('不支持分卷 ZIP');
  const count=b.readUInt16LE(end+10), size=b.readUInt32LE(end+12), offset=b.readUInt32LE(end+16);
  if (offset+size>end || count===65535) throw new Error('ZIP 目录边界无效');
  let p=offset;const entries=[];
  for (let n=0;n<count;n++) {
    if (p+46>offset+size || b.readUInt32LE(p)!==0x02014b50) throw new Error('ZIP 条目不完整');
    const nameSize=b.readUInt16LE(p+28),extra=b.readUInt16LE(p+30),comment=b.readUInt16LE(p+32);
    if(p+46+nameSize+extra+comment>offset+size) throw new Error('ZIP 名称越界');
    entries.push({name:b.subarray(p+46,p+46+nameSize).toString('utf8'),method:b.readUInt16LE(p+10),compressed:b.readUInt32LE(p+20),expanded:b.readUInt32LE(p+24)});
    p+=46+nameSize+extra+comment;
  }
  if(p!==offset+size) throw new Error('ZIP 目录长度不一致');
  return entries;
}
export function inspect(bytes) { return {members:listZipEntries(bytes)}; }
