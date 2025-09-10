# Indivdualized Exam Scripts
Scripts (written in Python) to help instructors generate individualized exams for students. Originally designed for a standards-based grading scheme, but should be flexible enough to use in a variety of circumstances. Let me know if you try to use this and it does/doesn't work for you. I'm open to suggestions to making it better!

For full transparency, I don't use this exact setup in my classes currently; I have the original version that I build up to work for my particular situation and needs. This should be a more general version that works fairly well, but let me know if there are issues you come across!

# Work is still in progress

ExamGeneration_v1 is the main file that builds the GUI and calls all other files.

## Parts that work
- Modified TeX file generation
- Individual Exam Generation
- Scanned Exam Processing

## To-Do
- Make everything work with paths instead of file names.
- Make the GUI look better
- Process into .exe file for easier use
- Make grade table options more clear and usable.


# Modified TeX file generation

This component is meant to take an exam file that you have already written and modify it to work with the remaining components. It will build the functions and macros needed to randomize the versions of the function across different exams and sections.

For the sectioning and versioning information, you must include the following:
- `\versNum` wherever you would like the version number to appear in the file
- `\secNum` wherever you want the section number(s) to appear
- `\stuName` in the name blank where students would normally write their name (it will be typed in if you have this part randomized)
You can use `\providecommand` lines at the start of the TeX file to allow you to still run the original exam before modifying it; see the SampleExam file for more details.

After this, you can typeset your exam like normal. Everything before the `\begin{questions}` will be included like normal. If you want to include a gradetable, use the plain `\gradetable` command. 

## Version and Section Specific Exams

If all you want to include is version and section specific exams, the only thing you need to do here is type all of the problems!
- You need to have the same number of versions of each problem that you want to shuffle around. (That is, you can't have two versions of the first problem and three of the second. You can duplicate the problems if you need to.)
- The problems of the same type need to be next to each other.
- When run, the code will generate one TeX file for each version/section pair that has a random collection of these problem versions on it.

## Individualized Exams
If you want to take advantage of the standards-based approach and individualize the exams for students, you need a few more things.
- Before each category of problem (each standard you are assessing), you need to provide a category title. This should be a somewhat short, but unique, alpha-numeric code for the problem type. This is given by putting the code above the question type as `%** CODE HERE **%`. Do not include any spaces between the asterisks and your code. Only put this once per group of questions you want to randomize.
- Anything you typeset between the code line and the first `\question` of that type will be included on every exam. You can use this for question labels or headers.
- You still need to have the same number of questions for each header.
- If you want to group categories (i.e., either all of them show up or none of them), you can give them the same code. You should repeat the code before each group of problems of the same type. The code will only show up once on the sheet where you specify which students get which problems, but it will have a number after it to indicate how many times it appeared.
- When run, this part of the code will output a spreadsheet with random student information on it. You should then clear out this information and fill in the appropriate data for your students before proceeding to the next part. 

# Individualized Exam Generation
This component will take the modified exam from the first component and the filled-in spreadsheet from the end of that process and write an individual PDF exam for each student.
- It will fill in their name, section, and version number where indicated.
- It will include exactly the problems marked Y on the spreadsheet, and will skip the ones marked N.
- Each exam will be stored as a separate PDF file for printing. 
- You have the option of getting the code to generate a blank exam of each type (useful for something like a Gradescope template), solutions (provided you wrote them in the first place), and empty directories for the scans (again, for use with Gradescope).

# Scanned Exam Processing
This wraps up the entire process of using something like Gradescope to grade individualized exam. This part of the code takes the scanned versions of the exams and adds in blank pages for each problem that was not included on that student's exam during generation. Therefore, each exam is the same full length again, and Gradescope can appropriately sort and organize them.

You will need to tell the program how many blank pages are at the start (including the title page) and end (for extra work) of the exam so that the exams are stepped through appropriately. Any number of exams can be included in each file, as long as there aren't extra or missing pages, it should work fine. This uses OCRmyPDF and Tesseract to process and read the files, so you'll need them and Ghostscript installed on your computer to run this.

You can experiment with the provided Sample Files. The `TrueRoster` table is meant to be the list of actual students, while `WrittenNames` has a few errors in it, which were used to generate the exams. The code will tell you when it can't find a student exactly and what it's choosing to do instead. 
