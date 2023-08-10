/*----- SQL Programming Handbook ------*/ 
/*      Author: Matthew Hoshauer
/*-------------------------------------*/


/* ----- Creating a Database --------*/

CREATE DATABASE name

USE name

DESCRIBE name

SELECT DATABASE(); <---- Shows list of currently available databases

CREATE TABLE pet (name VARCHAR(20), owner VARCHAR(20),
    species VARCHAR(20), sex CHAR(1), birth DATE, death DATE);