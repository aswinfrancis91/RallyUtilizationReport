# RallyUtilizationReport
A C# console app that I wrote in a hurry and uses Rally API to generate an email of team utilization data for an iteration


This application is by no means clean or the best but it gets the work done with the time I had

There are parts in the code that would need to be update with credentials and other info for the report to be pulled correctly. Email would be similar to this. From what I know, teams either use 'Actuals' or 'TimeSpent'. So in this this example, 'Actuals' is empty

| Iteration | Tasks | Capacity | Estimate | ToDo | Actuals | TimeSpent | Time/Cap % | Est/Cap % |
|-----------|-------|----------|----------|------|---------|-----------|------------|-----------|
| User 1    | 47    | 90.0     | 82.50    | 64   |         | 69.75     | 78         | 92        |
| User 2    | 53    | 90.0     | 98.5     | 33   |         | 73.5      | 82         | 109       |
