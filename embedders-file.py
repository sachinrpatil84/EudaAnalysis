import boto3
import json
import base64
from io import BytesIO
from PIL import Image
from config import AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY, AWS_REGION
from config import TEXT_EMBEDDER_MODEL_ID, IMAGE_EMBEDDER_MODEL_ID

class BedrockEmbedder:
    def __init__(self):
        """Initialize AWS Bedrock client"""
        self.bedrock_runtime = boto3.client(
            service_name='bedrock-runtime',
            region_name=AWS_REGION,
            aws_access_key_id=AWS_ACCESS_KEY_ID,
            aws_secret_access_key=AWS_SECRET_ACCESS_KEY
        )
    
    def get_text_embedding(self, text):
        """Get vector embedding for text using Amazon Titan Text Embedder"""
        try:
            # Prepare request body
            request_body = json.dumps({
                "inputText": text
            })
            
            # Call Bedrock API
            response = self.bedrock_runtime.invoke_model(
                modelId=TEXT_EMBEDDER_MODEL_ID,
                contentType='application/json',
                accept='application/json',
                body=request_body
            )
            
            # Parse response
            response_body = json.loads(response['body'].read())
            embedding = response_body.get('embedding', [])
            
            return embedding
        except Exception as e:
            print(f"Error getting text embedding: {str(e)}")
            return []
    
    def get_image_embedding(self, image_bytes):
        """Get vector embedding for image using Amazon Titan Image Embedder"""
        try:
            # Encode image to base64
            base64_image = base64.b64encode(image_bytes).decode('utf-8')
            
            # Prepare request body
            request_body = json.dumps({
                "inputImage": base64_image
            })
            
            # Call Bedrock API
            response = self.bedrock_runtime.invoke_model(
                modelId=IMAGE_EMBEDDER_MODEL_ID,
                contentType='application/json',
                accept='application/json',
                body=request_body
            )
            
            # Parse response
            response_body = json.loads(response['body'].read())
            embedding = response_body.get('embedding', [])
            
            return embedding
        except Exception as e:
            print(f"Error getting image embedding: {str(e)}")
            return []
