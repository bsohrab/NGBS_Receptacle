import boto3
import json
from collections import defaultdict


# Create SQS client
ACCESS_KEY = 'AKIATOBVOLOTY3CR4TMG'
SECRET_KEY = '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'
client = boto3.client('sqs', region_name='us-east-1', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)


queue_url = 'https://sqs.us-east-1.amazonaws.com/236335291303/testNGBS'


# Receive message from SQS queue
response = client.receive_message(
    QueueUrl=queue_url,
    AttributeNames=[
        'SentTimestamp'
    ],
    MaxNumberOfMessages=1,
    MessageAttributeNames=[
        'All'
    ],
    VisibilityTimeout=0,
    WaitTimeSeconds=0
)

try:
    message = response['Messages'][0]
    receipt_handle = message['ReceiptHandle']

    regulus = json.loads(response['Messages'][0]['Body'])["Records"][0]['s3']
    print(regulus)
    print(type(regulus))
    keyname = regulus['object']['key']
    bucketname = regulus['bucket']['name']
    print (keyname,'and', bucketname)
except:
    print("no more stuff")
client.delete_message(
    QueueUrl=queue_url,
    ReceiptHandle=receipt_handle
)
