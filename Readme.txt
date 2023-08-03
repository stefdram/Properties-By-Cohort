
Summary

Create a cohort analysis of company accounting data.
A cohort analysis groups property purchase by month, and see how it performs over the following months or time period
I have prepared a file propertiesByCohort.json with some data, and should be structured conveniently:

[
	{
		month: INT
		year: INT
		properties: [
			{
				propertyInfo..,
				payments: [
					{
						paymentInfo..
					},
					{
						paymentInfo..
					}
				],
			},
			{
				propertyInfo..,
				payments: [
					{
						paymentInfo..
					},
					{
						paymentInfo..
					}
				],
			}
		]
	}
]

Instructions:
1.  Import the json file and use JSON.parse() to convert into a typical object readable by NodeJS.
2.  Use XlsxJs (https://www.npmjs.com/package/xlsx) to create a cohort excel file (follow the cohort template)
3.  Recommend three functions:
	a.  createWorkbook
	b.  createWorksheetAndTemplate
	c.  fillWorkbook
