# Reference Intervals and Reference Codes

## Reference Intervals

referenceIntervals is a dictionary for numeric test results where the key is the LOINC test code and the value is the reference intervals data for that test.

referenceIntervals data consists of 'specimen' data and 'precision' data, where 'precesiion' is the number of decimal points to which the test data is recorded.

'specimen' data is dictionary where the key is the specimen type code or 'all' if the specimen data applies to all specimen types.
The value is a dictionary of reference intervals for each sex code where the sex code 'all' means that the reference intervals apply to any sex code.

For each sex code there will be a list of referenence intervals; one for each age range.

All reference intervals have a minimum age, maximum age, minimum value and maximum value, being the normal minimum value and normal maximum value for this age range. Minimum age and maximum age are coded as ISO-8601 formatted durations.  
e.g. "P\[0-9\]\[0-9\]*\[DWMY\]"

## Reference Codes for LOINC test codes

referenceCodes is a dictionary for non-numeric test results where the key is the LOINC test code and the value is the reference codes data for that test.

refereceCodes data is a dictionary where the key is the specimen type code or 'all' if the reference codes apply to all speciment types. The value is a dictionary of reference code sets for each sex code where the sex code 'all' means that the reference code set applies to any sex code.

For each sex code there will be a list of reference code sets; one for each age range.

All reference code sets have a minimum age, a maximum age and a list two element lists where the first element is the code and the second element the cumulative percentage of the population that matches this code, for this age range. Minimum age and maximum age are coded as ISO-8601 formatted durations.  
e.g. "P\[0-9\]\[0-9\]*\[DWMY\]"
