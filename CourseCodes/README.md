# Get your course codes
```JS
f=''
for(var i=2; i < $0.childNodes.length-2; i+=2) {
   f += $0.childNodes[i].cells[1].innerText + '\n'
}
console.log(f)
```