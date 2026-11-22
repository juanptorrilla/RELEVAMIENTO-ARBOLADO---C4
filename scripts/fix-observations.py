from pathlib import Path

path = Path("web/index.html")
source = path.read_text(encoding="utf-8")

old_field = '<textarea id="observaciones" placeholder="Observaciones del relevamiento..."></textarea>'
new_field = '<textarea id="observaciones" maxlength="1000" placeholder="Observaciones del relevamiento..."></textarea>'

old_pdf = '''   const obs=$("observaciones").value;
   if(obs){const lines=obs.match(/.{1,100}(?:\\s|$)/g)||[obs];lines.slice(0,5).forEach((line,i)=>putText(p1,line.trim(),100,707+i*22,s1,6.0));}'''
new_pdf = '''   const obs=$("observaciones").value.slice(0,1000);
   if(obs){
    const maxLines=5,maxWidth=580*s1,minSize=2.6;
    const wrapAtSize=(value,size)=>{
     const result=[];let line="";
     const fits=value=>font.widthOfTextAtSize(value,size)<=maxWidth;
     const words=value.replace(/\\s+/g," ").trim().split(" ").filter(Boolean);
     for(const word of words){
      if(!line){line=word; if(!fits(line)){let part="";for(const ch of word){if(!fits(part+ch)){if(part)result.push(part);part=ch;}else part+=ch;}line=part;while(result.length>=maxLines)return null;}continue;}
      if(fits(line+" "+word)){line+=" "+word;continue;}
      result.push(line);line="";
      if(result.length>=maxLines)return null;
      if(fits(word)){line=word;continue;}
      let part="";for(const ch of word){if(!fits(part+ch)){if(part)result.push(part);part=ch;if(result.length>=maxLines)return null;}else part+=ch;}line=part;
     }
     if(line)result.push(line);
     return result.length<=maxLines?result:null;
    };
    let lines=null,chosenSize=6;
    for(let size=6;size>=minSize;size-=0.2){const candidate=wrapAtSize(obs,size);if(candidate){lines=candidate;chosenSize=size;break;}}
    if(!lines)throw new Error("Las observaciones no entran en el espacio disponible. Reducí el texto.");
    lines.forEach((line,i)=>putText(p1,line,100,707+i*9,s1,chosenSize));
   }'''

if old_field not in source:
    raise SystemExit("No se encontró el campo de observaciones esperado; no se modificó el archivo.")
if old_pdf not in source:
    raise SystemExit("No se encontró el bloque de generación de observaciones esperado; no se modificó el archivo.")
source = source.replace(old_field, new_field, 1).replace(old_pdf, new_pdf, 1)
path.write_text(source, encoding="utf-8")
print("Observaciones: límite de 1.000 caracteres y ajuste dinámico en cinco renglones aplicado.")
