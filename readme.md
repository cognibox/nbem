# Non birthday Employee Matching (NBEM)

Tool to match employee with or without no birthday from contractors employees in CBX

From Microsoft Powershell use the following (requires Docker)
> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t nbem -q https://github.com/cognibox/nbem.git) <cbx_employee_db_dump.xlsx> <hc_list.xlsx> <results.xlsx>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t nbem -q https://github.com/cognibox/nbem.git) -h
