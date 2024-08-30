from __future__ import print_function
import time
import cloudmersive_convert_api_client
from cloudmersive_convert_api_client.rest import ApiException
from pprint import pprint
# Configure API key authorization: Apikey
configuration = cloudmersive_convert_api_client.Configuration()
configuration.api_key['Apikey'] = 'YOUR_API_KEY'
# Uncomment below to setup prefix (e.g. Bearer) for API key, if needed
# configuration.api_key_prefix['Apikey'] = 'Bearer'
# create an instance of the API class
api_instance = cloudmersive_convert_api_client.ConvertDocumentApi(cloudmersive_convert_api_client.ApiClient(configuration))
input_file = '/path/to/file' # file | Input file to perform the operation on.
body_only = true # bool | Optional; If true, the HTML string will only include the body of the MSG. Other information such as subject will still be given as properties in the response object. Default is false. (optional)
include_attachments = true # bool | Optional; If false, the response object will not include any attachment files from the input file. Default is true. (optional)
try:
# Convert Email MSG file to HTML string
api_response = api_instance.convert_document_msg_to_html(input_file, body_only=body_only, include_attachments=include_attachments)
pprint(api_response)
except ApiException as e:
print("Exception when calling ConvertDocumentApi->convert_document_msg_to_html: %s\n" % e)