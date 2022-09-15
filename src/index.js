#!/usr/bin/env node

import chalk from 'chalk';
import fs from 'fs';
import path from 'path';
import parse from 'csv-parse'

function isoDateStr(d) {
    return d.toISOString().split('T')[0];
}

async function getClasslist(dirPath=".") {

    var classlist = new Set();

    const filePath = path.join(dirPath, 'classlist.csv');

    if ( ! fs.existsSync(filePath) ) {
        return classlist;
    }

    const parser = fs
                    .createReadStream(filePath)
                    .pipe(parse({ 
                        delimiter: ',',
                        bom: true
                    }));

    let firstLine = true;
    let fnameIdx, lnameIdx;
    for await ( const record of parser ) {

        if ( firstLine ) {
            let i = 0;
            for ( const field of record ) {
                if ( field == "Last Name" ) {
                    lnameIdx = i;
                } else if ( field == "First Name" ) {
                    fnameIdx = i;
                }
                i += 1; 
            }
            firstLine = false;
        } else {

            const fname = record[fnameIdx].replace(/\s/g, "").toUpperCase();
            const lname = record[lnameIdx].replace(/\s/g, "").toUpperCase();
            classlist.add(`${fname} ${lname}`);
        }
    }

    return classlist;
}

async function getAttendanceFilepaths(dirPath=".") {
    var names = await fs.promises.readdir(dirPath);
    return names.filter(filename => { return filename.startsWith("meetingAttendance"); })
                .map(filename => path.join(dirPath, filename));
}

async function getAttendanceDataFromAttendanceReport(filePath) {

    const parser = fs.createReadStream(filePath, {encoding: 'utf-16le'})
                        .pipe(parse({ 
                            delimiter: '\t',
                            bom: true,
                            columns: ["Full Name", "Join Time", "Leave Time", "Duration", "Email", "Role", "Participant ID (UPN)"],
                            skip_records_with_error: true,
                            from_line: 9
                        }));

    let sessionDate;

    let studentsInAttendance = new Set();

    for await ( const record of parser ) {
        const date = new Date(record['Join Time'].substr(0, record['Join Time'].indexOf(",")));
        if ( sessionDate && sessionDate.getTime() !== date.getTime()) {
            console.warn(`Inconsistent session dates in ${filePath}`)
        } else {
            sessionDate = date;
        }

        studentsInAttendance.add(record['Full Name'].toUpperCase());  // Student name
    }

    return { date: isoDateStr(sessionDate), studentsInAttendance }
}
async function getAttendanceDataFromAttendanceList(filePath) {

    const parser = fs.createReadStream(filePath, {encoding: 'utf-16le'})
                        .pipe(parse({ 
                            delimiter: '\t', 
                            from_line: 2,
                            bom: true,
                            columns: ["Full Name", "User Action", "Timestamp"],
                            skip_records_with_error: true
                        }));

    let sessionDate;

    let studentsInAttendance = new Set();

    for await ( const record of parser ) {
        const date = new Date(record['Timestamp'].substr(0, record['Timestamp'].indexOf(",")));
        if ( sessionDate && sessionDate.getTime() !== date.getTime()) {
            console.warn(`Inconsistent session dates in ${filePath}`)
        } else {
            sessionDate = date;
        }

        studentsInAttendance.add(record['Full Name'].toUpperCase());  // Student name
    }

    return { date: isoDateStr(sessionDate), studentsInAttendance }
}

async function getAttendanceRecord(attendanceFilepaths, studentsEnrolled) {

    const data = [];
    let studentsInAttendance = new Set();
    for ( const attendanceFilepath of attendanceFilepaths ) {
        
        // There are two kinds of attendance reports (WTF MS Teams?)
        // Here, we determine which one we're working with
        const fileContent = fs.readFileSync(attendanceFilepath, { encoding: 'utf-16le' }).toString();
        let sessionData;
        if ( fileContent.match(/Meeting Summary/) ) {
            sessionData = await getAttendanceDataFromAttendanceReport(attendanceFilepath);
        } else {
            sessionData = await getAttendanceDataFromAttendanceList(attendanceFilepath);
        }

        // Take the union of the current set of students and the students who were in this session
        for ( const student of sessionData.studentsInAttendance ) {
            studentsInAttendance.add(student);
        }

        data.push(sessionData);
    }

    data.sort( (s1, s2) => {
        if ( s1.date === s2.date ) return 0;
        return s1.date > s2.date ? 1 : -1;
    })

    return { studentsEnrolled, studentsInAttendance, sessions: data };
}

function byLastName(a,b) { 
    if ( a.lname == b.lname ) {
        if ( a.fname == b.fname ) { return 0; }
        return a.fname > b.fname ? 1 : -1;
    } 
    return a.lname > b.lname ? 1 : -1;
}
function byFirstName(a,b) { 
    if ( a.fname == b.fname ) {
        if ( a.lname == b.lname ) { return 0; }
        return a.lname > b.lname ? 1 : -1;
    } 
    return a.fname > b.fname ? 1 : -1;
}
function byRatio(a,b) { 
    const ratioA = eval(a.ratio);
    const ratioB = eval(b.ratio);
    if ( ratioA == ratioB ) { return 0; } 
    return ratioA > ratioB ? 1 : -1;
}

function showAttendanceReport(attendanceRecord) {

    function showSingleRecord({fname, lname, record, ratio}, longestName) {
        const pad = longestName - fname.length - lname.length;

        const r = eval(ratio);
        let ratioStr;
        if ( r > .95 ) {
            ratioStr = chalk.greenBright(ratio);
        } else if  ( r > .66 ) {
            ratioStr = chalk.yellow(ratio);
        } else if ( r > .33 ) {
            ratioStr = chalk.magenta(ratio);
        } else {
            ratioStr = chalk.red(ratio);
        }

        let recordStr = record.map(r => r == 'P' ? chalk.greenBright(r) : chalk.redBright(r)).join("");

        console.log( `${' '.repeat(pad)}${lname}, ${fname} : ${ratioStr} ${recordStr} `);
    }

    function showSessionHeadings(sessions, leftPad=0) {
        const letters = "JFMAMJYASOND";
        let monthHeading = sessions.reduce((acc, s) => acc + letters[parseInt(s.date.split("-")[1])-1], "");
        console.log(" ".repeat(leftPad) + monthHeading);
        let tensHeading = sessions.reduce((acc, s) => acc + s.date[8], "");
        console.log(" ".repeat(leftPad) + tensHeading);
        let onesHeading = sessions.reduce((acc, s) => acc + s.date[9], "");
        console.log(" ".repeat(leftPad) + onesHeading);
        

        console.log(" ".repeat(leftPad) + "-".repeat(sessions.length));
    }

    let enrolledData = [];
    let unenrolledData = [];
    const numSessions = attendanceRecord.sessions.length;
    let longestName = 0;
    for ( const student of attendanceRecord.studentsInAttendance ) {
        let studentRecord = [];
        let timesPresent = 0;
        for ( const session of attendanceRecord.sessions) {
            if ( session.studentsInAttendance.has(student) ) {
                studentRecord.push("P");
                timesPresent += 1;
            } else {
                studentRecord.push("A");
            }
        }
        
        let data = { 
            fname: student.split(" ")[0],
            lname: student.split(" ")[1],
            record: studentRecord,
            ratio: `${timesPresent}/${numSessions}`
        };

        if ( attendanceRecord.studentsEnrolled.has(student) ) {
            enrolledData.push(data);
        } else {
            unenrolledData.push(data);
        }

        if ( student.length > longestName ) {
            longestName = student.length;
        }
    }

    // Handle enrolled students that have never been present
    const enrolled = attendanceRecord.studentsEnrolled
    const attended = attendanceRecord.studentsInAttendance
    // Set difference: enrolled - attended
    const absentees = new Set(
        [...enrolled].filter( s => !attended.has(s) )
    )
    absentees.forEach( s => enrolledData.push({ 
        fname: s.split(" ")[0],
        lname: s.split(" ")[1],
        record: Array(numSessions).fill("A"),
        ratio: `0/${numSessions}`
    }))

    console.log(chalk.gray("Sessions:"));
    for ( const session of attendanceRecord.sessions) {
        console.log(chalk.gray(session.date));
    }
    
    if ( enrolledData.length ) {
        console.log();
        console.log(chalk.gray(`Enrolled Attendees (${enrolledData.length}):`));
        showSessionHeadings(attendanceRecord.sessions, longestName+9);
        for ( const record of enrolledData.sort(byRatio) ) {
            showSingleRecord(record, longestName);
        }
        console.log();
        console.log(chalk.gray("Unenrolled Attendees:"));
    }

    showSessionHeadings(attendanceRecord.sessions, longestName+9);
    for ( const record of unenrolledData.sort(byRatio) ) {
        showSingleRecord(record, longestName);
    }

}

if ( process.argv.length !== 3 ) {
    console.log("Analyze My Attendance Records (amar)");
    console.log("Version 1.0.2");
    console.log();
    console.log("Usage:");
    console.log("amar <path>");
    console.log("");
    console.log("<path> must a path to a directory containing MS Teams attendance reports with names starting with 'meetingAttendance'")
} else {
    var dirPath = process.argv.slice(2)[0];

    if ( ! fs.existsSync(dirPath) ) {
        console.log(`The path ${dirPath} does not exist`);
    } else {

        const studentsEnrolled = await getClasslist(dirPath);
        const attendanceFilePaths = await getAttendanceFilepaths(dirPath)
        let attendanceRecord = await getAttendanceRecord(attendanceFilePaths, studentsEnrolled);
    
        showAttendanceReport(attendanceRecord);
    }

}
