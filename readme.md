# Non birthday Employee Matching (NBEM)

Tool to match employee with no birhtday from contractors employees in CBX

From Microsoft Powershell use the following (requires Docker)
> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/nbem.git) <cbx_contractor_db_dump.csv> <hc_list.xlsx> <results.xlsx>

To see the command line tool help use the following:

> docker run --rm -it -v ${pwd}:/home/script/data $(docker build -t icm -q https://github.com/iguzu/nbem.git) -h