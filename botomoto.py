import boto3
import reviewTool2 as rt
from boto3.session import Session



ACCESS_KEY = 'AKIATOBVOLOTY3CR4TMG'
SECRET_KEY = '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'
session = Session(aws_access_key_id=ACCESS_KEY,
              aws_secret_access_key=SECRET_KEY)
localLocation = "C:\\Users\\spasikhani\\Downloads\\"




s3 = session.resource('s3')

#download and convert
targetbucket = s3.Bucket('ngbsbucket')
for s3_file in targetbucket.objects.all():
    print(s3_file.key.split("/")[0]) # prints the contents of bucket
    localLocation = "C:/Users/spasikhani/Desktop/venv/drop folder/"+ s3_file.key.split("/")[0]
    targetbucket.download_file(s3_file.key, localLocation)
    #rt.task(localLocation)
#find and reupload



#send message to AWS that the
client = boto3.client('sqs', region_name='us-east-1', aws_access_key_id=ACCESS_KEY, aws_secret_access_key=SECRET_KEY)
response = client.send_message(
    QueueUrl='https://sqs.us-east-1.amazonaws.com/236335291303/testNGBS',
    MessageBody='The file download is done',
    DelaySeconds=12)