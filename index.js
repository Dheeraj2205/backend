/*
// // console.log("Hello world !!")
// // let number = 1; 
// // number = 2;

// // const myFunction = () => {
// //     //Logic goes here
// //     console.log("From myFunction")
// // }

// // myFunction();
// // console.log(number)
// const studentObject = {
//     NAME : "Dheeraj",
//     CGPA : 0,
//     address : {
//         city : "Jamshedpur",
//         state : "Jahrkhand",
//         region : "Telco"
//     },
//     "Favourite Hobby": "Watching Sports"
// }

// console.log(studentObject.NAME)
// console.log(studentObject.CGPA)
// console.log(studentObject.address.city)
// console.log(studentObject["Favourite Hobby"])
// // console.log(studentObject.address.state)
// // console.log(studentObject.address.region)

// const { NAME } = studentObject
// console.log(NAME)
*/
const parser = require('simple-excel-to-json')
const doc = parser.parseXls2Json('./Example.xlsx')[0]; 
// const json2xls = require("json2xls")
const fs = require("fs");
const json2xls = require('json2xls')

console.log(doc)

//map, reduce, filter :

const totlacgpa = doc.reduce((prevValue, currentValue) => {
    // console.log(currentValue)
    prevValue += currentValue.CGPA
    return prevValue;
}, 0)

const averageCgpa = totlacgpa / doc.length;

const gradeDocument = doc.map(student => {
    if (student.CGPA > 9.5) {
        student.GRADE = "A+"
    }
    else if (student.CGPA <= 9.5 && student.CGPA > 9)    {
        student.GRADE = "A"
    }
    else if (student.CGPA <= 9 && student.CGPA > 8.5)    {
        student.GRADE = "B"
    }
    else if (student.CGPA <= 8.5 && student.CGPA > 7)    {
        student.GRADE = "C"
    }
    else if (student.CGPA <= 7 && student.CGPA > 5)    {
        student.GRADE = "D"
    }
    else if (student.CGPA <= 5 && student.CGPA > 3)    {
        student.GRADE = "E"
    }
    else {
        student.GRADE = "F"
    }
    return student;
} )

console.log(totlacgpa)
console.log(averageCgpa)
const filterDocument = gradeDocument.filter(student => student.fs)
gradeDocument.push({CGPA : averageCgpa, NAME : "Average Grades" })
const excelDocument = json2xls(gradeDocument)

fs.writeFileSync("Grades.xlsx", excelDocument, "binary")
console.log(gradeDocument)
console.log(filterDocument) 

