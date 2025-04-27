import cv2
import numpy as np
import fitz
import os

class ImageMatcher:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path

    def extract_and_match_images(self, target_image_path):
        threshold = 10  # Threshold for image comparison
        doc = fitz.open(self.pdf_path)
        target_image = cv2.imread(target_image_path)

        if target_image is None:
            print("Error in fetching the image (invalid image)")
            return None

        found_logo_page = None

        for page_num, page in enumerate(doc, start=1):
            for img_index, img in enumerate(page.get_images(full=True), start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)  # Load as color image

                # Compare the extracted image with the target image
                mse = self.compare_images(target_image, image)
                if mse < threshold:
                    found_logo_page = page_num
                    break  # No need to continue checking images on this page if logo is found
        
            if found_logo_page is not None:
                break

        return found_logo_page

    def compare_images(self, image1, image2):
        # Resize image1 to match the dimensions of image2
        image1_resized = cv2.resize(image1, (image2.shape[1], image2.shape[0]))

        # Compute Mean Squared Error (MSE)
        mse = np.mean((image1_resized - image2) ** 2)
        return mse

class PDFImageMatcher(ImageMatcher):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)

    def compare_text_and_image(self, provided_account_value, account_category_value):
        """
        Compare the extracted text with the provided account code, and match images if necessary.
        """
        # Coordinates of the region containing the account number
        if account_category_value == "B2B":
            x1 = 290.0
            y1 = 119.07999267599996
            x2 = 335.44
            y2 = 127.07999267599996
        elif account_category_value == "B2C":
            x1 = 290.0  # Example coordinates for category 2
            y1 = 119.07999267599996
            x2 = 335.44
            y2 = 127.07999267599996
        else:
            return False, None  # Return False if category is invalid

        extracted_text = ""  # Variable to store the extracted text
        doc = fitz.open(self.pdf_path)
        logo_present_on_all_pages = True  # Flag to track if logo is present on all pages
        pages_without_logo = []  # List to store pages without the logo

        # Load the target logo
        target_logo_path = "logo_business.jpeg" if account_category_value == "B2B" else "logo.jpeg"
        target_logo_path = os.path.join("promo_logo_images", target_logo_path)
        target_logo = cv2.imread(target_logo_path)

        # Iterate over each page to check text and logo
        last_page_blank = False  # Flag to track if the last page is blank
        for page_num, page in enumerate(doc, start=1):
            if page_num == len(doc):  # Check if it's the last page
                last_page_blank = page.get_text().strip() == ""  # Check if the last page is blank
                if last_page_blank:
                    break  # If last page is blank, skip logo verification

            region_text = page.get_text("text", clip=(x1, y1, x2, y2)).strip()
            extracted_text += region_text + "\n"

            # Check if the extracted text matches the provided account code
            if provided_account_value.lower() not in extracted_text.lower():
                print("Account number not found.")
                return False, None  # Return False if account number not found

            # Check if logo is present on the current page
            found_logo = False
            for img_index, img in enumerate(page.get_images(full=True), start=1):
                xref = img[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                extracted_logo = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)

                # Compare the extracted logo with the target logo
                mse = self.compare_images(target_logo, extracted_logo)
                if mse < 10:  # Adjust threshold as needed
                    found_logo = True
                    break

            if not found_logo:
                logo_present_on_all_pages = False
                pages_without_logo.append(page_num)

        # If logo verification is required and the last page is blank, skip logo verification
        if not last_page_blank:
            # If logo is missing on any page, raise an error
            if not logo_present_on_all_pages:
                raise ValueError(f"Logo not found on pages: {', '.join(map(str, pages_without_logo))}")

        # All conditions met, return True
        return True, page_num

class PromoMatcher(ImageMatcher):
    def __init__(self, pdf_path):
        super().__init__(pdf_path)

    def match_promo(self, account_number, market_value, account_category_value, language_value, promo_image):
        """
        Match the provided image with the image extracted from the PDF.
        """
        # Extract and match images
        return self.extract_and_match_images(promo_image)

def generate_image_filename(account_category_value, market_value, language_value):
    return f"{account_category_value}_PROMO_{market_value}_{language_value}.jpg"

def process_image(image_filename):
    if image_filename == "Invalid parameter":
        return False
    if image_filename == "":
        return False
    # Your code to process the image goes here
    # Replace the return statement with your processing logic
    # Return True if processing is successful, False otherwise
    return True  # Or return False based on your actual processing logic

def extract_specific_images(pdf_path, save_path, page_number, image_indices, account_number):
    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    doc = fitz.open(pdf_path)
    
    for page_num, page in enumerate(doc, start=1):
        if page_num != page_number:
            continue  # Skip pages until the desired page is reached
            
        for img_index, img in enumerate(page.get_images(full=True), start=1):
            if img_index not in image_indices:
                continue  # Skip images until the desired ones are reached
            
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), -1)
            
            # Save the image with account number, page number, and image number in the filename
            image_save_path = os.path.join(save_path, f"{account_number}_page_{page_num}_image_{img_index}.png")
            cv2.imwrite(image_save_path, image)
            print(f"Image saved: {image_save_path}")

def match_logo(pdf_path, account_number, account_category_value):
    try:
        logo_matcher = PDFImageMatcher(pdf_path)
        return logo_matcher.compare_text_and_image(account_number, account_category_value)
    except ValueError as e:
        print(e)  # Print the error message
        # Get the last accessed page number
        if hasattr(e, 'args') and len(e.args) > 0 and isinstance(e.args[0], int):
            logo_page = e.args[0]
        else:
            logo_page = None

        if logo_page is not None:
            print(f"Logo not matched on page {logo_page}.")
       
def match_promo(pdf_path, account_number, market_value, account_category_value, language_value):
    promo_image = generate_image_filename(account_category_value, market_value, language_value)
    if promo_image == "Invalid parameter":
        print("Invalid parameter provided. Aborting image verification.")
        return False, None
    promo_image = os.path.join("promo_logo_images",promo_image)
    promo_matcher = PromoMatcher(pdf_path)
    return promo_matcher.match_promo(account_number, market_value, account_category_value, language_value, promo_image)

# # Usage
# pdf_path = "test123.pdf"
# account_number = "99002460"
# market_value = "SXM"
# account_category_value = "B2B"
# language_value = "DUT"
# save_path = "extracted_images"
# page_number = 1
# image_indices = [2, 3]  # Specify the indices of the images you want to extract
# extract_specific_images(pdf_path, save_path, page_number, image_indices, account_number)

# promo_image = generate_image_filename(account_category_value, market_value, language_value)

# if promo_image == "Invalid parameter":
#     print("Invalid parameter provided. Aborting image verification.")
# else:
#     # First, let's match the logo
#     logo_matcher = PDFImageMatcher(pdf_path)
#     try:
#         logo_result, logo_page = logo_matcher.compare_text_and_image(account_number, account_category_value)
#         if logo_result:
#             print(f"Logo matched on all pages!")
#     except ValueError as e:
#         print(e)  # Print the error message
#         # Get the last accessed page number
#         if hasattr(e, 'args') and len(e.args) > 0 and isinstance(e.args[0], int):
#             logo_page = e.args[0]
#         else:
#             logo_page = None

#         if logo_page is not None:
#             print(f"Logo not matched on page {logo_page}.")

#     # Now, let's match the promo image
#     promo_matcher = PromoMatcher(pdf_path)
#     promo_page = promo_matcher.match_promo(account_number, market_value, account_category_value, language_value, promo_image)

#     if promo_page is not None:
#         print(f"Promo matched on page {promo_page}!")
#     else:
#         print("Promo not matched.")
