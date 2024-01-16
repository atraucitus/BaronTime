# ***BaronTime***
*Converts your course excel timetable to a nice calendar. Can be used for Office and Google Calendar*
[Get your course codes from (Bennett DigiiCampus)](#get-your-course-codes)


## How to:
- run ```pip install -r requirements.txt```
- Run the ```main.py``` file without any parameters.
- Select the correct excel file with your timetable.
- Follow prompts
    - Ensure your courses exist in 

## Get your course codes

Generate your course registration slip from [CollPoll](https://bennett.digiicampus.com/courseRegistration/student).

Slip looks like this
![Course Registration](image.png)

Open Dev Tools and run the following code in the console window.
```JS
// Amatuer JS Codes.
table = document.childNodes[1].childNodes[1].childNodes[1].childNodes[0].childNodes[0].childNodes[4].childNodes[2].childNodes[9].childNodes[1].childNodes[1]

f=''
for(var i=2; i < table.childNodes.length-2; i+=2) {
   f += table.childNodes[i].cells[1].innerText + '\n'
}
console.log(f)
```
