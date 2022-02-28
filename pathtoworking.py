import boto3
import reviewTool2 as rt
from boto3.session import Session
import boto3
import json
import os
from collections import defaultdict

################################################################
## Create SQS client for incoming files to translate
ACCESS_KEY = 'AKIATOBVOLOTY3CR4TMG'
SECRET_KEY = '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'
queue_url_recieve = 'https://sqs.us-east-1.amazonaws.com/236335291303/testNGBS'
queue_url_outgoing = 'https://sqs.us-east-1.amazonaws.com/236335291303/RecieptQ'
localLocation = "C:\\Users\\spasikhani\\Desktop\\venv\\drop folder\\"
client = boto3.client('sqs', region_name='us-east-1', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)





################################################################
## Receive message from SQS queue
#and parse it for information on the bucket and key
#of the file that was just updated
response = client.receive_message(
    QueueUrl=queue_url_recieve,
    AttributeNames=[
        'SentTimestamp'
    ],
    MaxNumberOfMessages=10,
    MessageAttributeNames=[
        'All'
    ],
    VisibilityTimeout=0,
    WaitTimeSeconds=0
)
message = response['Messages'][0]
receipt_handle = message['ReceiptHandle']
try:
    regulus = json.loads(response['Messages'][0]['Body'])["Records"][0]['s3']
    keyname = regulus['object']['key']
    bucketname = regulus['bucket']['name']
    print (keyname,'and', bucketname)
    print(regulus)
except:
    print("no more stuff")
    #delete the message after recieving it


################################################################
##Go into the s3 bucket and
#use the key to download, translate,
#and reupload the file back to the containing directory
session = Session(aws_access_key_id=ACCESS_KEY,aws_secret_access_key=SECRET_KEY)
s3 = session.resource('s3')
#download and convert
targetbucket = s3.Bucket(bucketname)
print("working")
localLocation = localLocation + keyname
targetbucket.download_file(keyname, localLocation)#, localLocation)
print("working")
keyname = rt.task(localLocation)
#find and reupload
response = targetbucket.upload_file(keyname,"oolongjohnson.xlsx")
print("reentry successful")


################################################################
## if the tasks are completed successfully, delete the message
# and in the event that the tasks are not completed successfully,
#send a message saying to redo or the issue that went wrong
client.delete_message(
    QueueUrl=queue_url_recieve,
    ReceiptHandle=receipt_handle
)





################################################################
## After deleting the message
#send message to AWS that the file has been uploaded as to trigger the lambda function
#to parse the message for the bucket and key needed to get the file
#and move it to the customer portal into the documents section
client = boto3.client('sqs', region_name='us-east-1', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)
response = client.send_message(
    ##need a new queue for the return messages until i figure out more
    QueueUrl=queue_url_outgoing,
    ##convert this message bodTY from bad string to chad dictionary json parseable format, 1 layer.
    MessageBody='The following item has been completed ' + keyname+ ' in bucket ' + bucketname,
    ##this doesent matter
    DelaySeconds=12)