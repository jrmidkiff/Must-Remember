# Must-Remember
This is code for things that I have to remember how to do for work that I am otherwise quite likely to forget


## SQL  
Using CASE (equivalent to switch or multiple if-elses)  

    SELECT COUNT(*), 
    CASE 
        WHEN number_grade > 90 THEN 'A'
        WHEN number_grade > 80 THEN 'B'
        WHEN number_grade > 70 THEN 'C'
        ELSE 'F'
        END AS 'letter_grade'
    FROM student_grades
    GROUP BY letter_grade;
