from io import BytesIO
from typing import Annotated
from semantic_kernel.functions.kernel_function_decorator import kernel_function
from PIL import Image
import json
from openai import AzureOpenAI
import dotenv
import os

class ImageGenerator:
    """The Image Generator draws images to add to the presentation contents."""

    @kernel_function(description="Draw and return an image.")
    def draw_image(self, 
                            description: Annotated[str, "The description of the image to be drawn"],
                            ) -> Annotated[str, "Draw and return an image."]:

        print("Generating image...")

        # Load environment variables from .env file
        dotenv.load_dotenv('../.env')

        AZURE_OPENAI_DALLE_ENDPOINT = os.getenv("AZURE_OPENAI_DALLE_ENDPOINT")
        AZURE_OPENAI_DALLE_KEY=os.getenv("AZURE_OPENAI_DALLE_KEY")
        AZURE_OPENAI_DALLE_DEPLOYMENT=os.getenv("AZURE_OPENAI_DALLE_DEPLOYMENT")
        # Note: DALL-E 3 requires version 1.0.0 of the openai-python library or later

        #chat bot tells the user that it is drawing an image based on the prompt
        print(f"Drawing an image: {description}")

        client = AzureOpenAI(
            api_version="2024-02-01",
            azure_endpoint=AZURE_OPENAI_DALLE_ENDPOINT,
            api_key=AZURE_OPENAI_DALLE_KEY,
        )

        try:
            result = client.images.generate(
                model=AZURE_OPENAI_DALLE_DEPLOYMENT, # the name of your DALL-E 3 deployment
                prompt=description, 
                n=1
            )
        except Exception as e:
            print(f"An error occurred: {e}")
            return "An error occurred while generating the image."

        image_url = json.loads(result.model_dump_json())['data'][0]['url']
        print(f"Image URL: {image_url}")
        return image_url

