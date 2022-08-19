download boto3 library into package folder
```
cd ~/python-customers-lambda/AddCustomer
pip3 install -t package boto3
```

zip package folder
```
cd package
zip -r ../pythonaddCustomerFunction.zip . 
```

add addCustomer.py to the zip file
```
cd ..
zip -g pythonaddCustomerFunction.zip addCustomer.py
```

deploy lambda create-function via zipfile
```
aws lambda create-function --function-name AddCustomer \
--runtime python3.9  \
--role $ROLE_ARN_READWRITE \
--handler addCustomer.addCustomerToFile \
--publish \
--zip-file fileb://pythonaddCustomerFunction.zip
```

invoke lambda function
```
aws lambda invoke --function-name AddCustomer --payload fileb://newCustomerPayload.json output.txt ; cat output.txt
```
