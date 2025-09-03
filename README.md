# Project Title
Java-powered Powerpoint Generation

## Description
 A spring based Java application that uses the Apache Poi API for reading data from an Excel file and creating a Powerpoint from that data using the Aspose file format API.
 The application reads project data from an Excel file and generates a Powerpoint with a dashboard with a tabular counterpart. 
 It is a work-in-progress with the current version generating the initial dashboard slide with the tabular project data. Later iterations will add summary charts and slides for each of the individual projects.

## Features
- Read tabular excel source data using the Apache POI API.
- Generate a Powerpoint presentation using the Aspose API to create a Dashboard slide, and individual detail slides for each project
- Use Aspose API for styling the slides, adding color, adding images, and generating charts.

## Getting Started
- Java 11
- Maven 
- Application directory structure (Windows profile):
  -- C:\DemoApp\bin (executable jar file)
  -- C:\DemoApp\conf\ (excel source configurations.xlsx)
  -- C:\DemoApp\output\ (generated powerpoint location)
  -- C:\DemoApp\images\ (images location)
 - Note that the pom.xml for this project has currently only 1 'dev' profile in the build section, which caters to the Windows directory structure. You can add another profile for a directory structure for a different OS.

## Execute Application Command
java -jar powerpoint-creator-1.0-jar-with-dependencies.jar (from insude bin folder)
