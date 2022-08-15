echo "URL: https://$MYBUCKET.s3.amazonaws.com/dragonsapp/index.html"

https://sodfzfj9qa.execute-api.us-east-1.amazonaws.com/prod/dragons

https://sodfzfj9qa.execute-api.us-east-1.amazonaws.com/prod/dragons?family=red
/prod/dragons?family=red
https://sodfzfj9qa.execute-api.us-east-1.amazonaws.com/prod/dragons?family=blue
/prod/dragons?family=blue
/prod/dragons?dragonName=Atlas

post 
req body:
validate ok 
{
  "dragonName": "Frank",
  "description": "This dragon is brand new, we don't know much about it yet.",
  "family": "purple",
  "city": "Seattle",
  "country": "USA",
  "state": "WA",
  "neighborhood": "Downtown",
  "reportingPhoneNumber": "15555555555",
  "confirmationRequired": false
}

validate fail 
{
  "dragonName": "Frank",
  "description": "This dragon is brand new, we don't know much about it yet.",
  "family": "purple",
  "city": "Seattle",
  "country": "USA",
  "state": "WA",
  "neighborhood": "Downtown",
  "reportingPhoneNumber": "15555555555",
  "confirmationRequired": "no, thank you"
}

echo "URL: https://$MYBUCKET.s3.amazonaws.com/dragonsapp/index.html"
URL: https://mybucketnancy.s3.amazonaws.com/dragonsapp/index.html

