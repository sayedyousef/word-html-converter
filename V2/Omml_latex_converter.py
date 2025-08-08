#Omml_latex_converter.py
import os
import lxml.etree as etree
from mathml2latex import convert

# NOTE: You MUST provide the path to the 'OMML2MML.XSL' stylesheet.
# This file is typically located in your Microsoft Office installation directory.
# For example, on a 64-bit Windows machine with Office 16, the path might be:
# "C:\\Program Files\\Microsoft Office\\root\\Office16\\OMML2MML.XSL"
# You may need to copy this file to a location your script can access.
XSL_STYLESHEET_PATH = "C:\\path\\to\\OMML2MML.XSL"

# --- Main conversion logic ---

class OmmlConverter:
    """
    A utility class to convert OMML to LaTeX using an XSLT stylesheet
    and a pure Python MathML to LaTeX converter.
    """

    def __init__(self, xslt_path):
        """
        Initializes the converter by pre-loading the XSLT stylesheet.

        Args:
            xslt_path (str): The file path to the OMML2MML.XSL stylesheet.
        """
        if not os.path.exists(xslt_path):
            raise FileNotFoundError(
                f"The XSLT stylesheet was not found at the specified path: {xslt_path}"
            )
        
        # Parse the XSLT stylesheet once for performance
        xslt_root = etree.parse(xslt_path)
        self.transform = etree.XSLT(xslt_root)

    def _omml_to_mathml(self, omml_string):
        """
        Converts an OMML string to a MathML string using the pre-loaded XSLT.

        Args:
            omml_string (str): The raw OMML XML string.

        Returns:
            str: The converted MathML XML string.
        """
        try:
            omml_root = etree.fromstring(omml_string.encode('utf-8'))
            mathml_root = self.transform(omml_root)
            return str(mathml_root)
        except Exception as e:
            # Catch all exceptions and return a placeholder for the caller
            print(f"Error during OMML to MathML conversion: {e}")
            return None

    def convert_omml_to_latex(self, omml_string):
        """
        Converts an OMML string to a LaTeX string using a two-step process.

        This method mirrors the signature of your other extraction methods.

        Args:
            omml_string (str): The input OMML content as a string.

        Returns:
            str: The converted LaTeX string, or a placeholder if an error occurs.
        """
        try:
            # Step 1: Convert OMML to MathML using XSLT
            mathml_string = self._omml_to_mathml(omml_string)
            if not mathml_string:
                return "[equation]"

            # Step 2: Convert MathML to LaTeX using the mathml2latex library
            latex_string = convert(mathml_string)
            return latex_string
        except Exception as e:
            # Catch any conversion errors and return a placeholder
            print(f"Error during MathML to LaTeX conversion: {e}")
            return "[equation]"


# --- Example Usage ---

if __name__ == "__main__":
    # Sample OMML for a simple quadratic formula:
    # x = (-b ± sqrt(b^2 - 4ac)) / (2a)
    sample_omml = """
    <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/officeDocument/2006/wordml">
        <m:eqArr>
            <m:eq>
                <m:oMath>
                    <m:r><m:t>x</m:t></m:r>
                    <m:r><m:t xml:space="preserve"> = </m:t></m:r>
                    <m:f><m:den><m:r><m:t>2a</m:t></m:r></m:den>
                        <m:num>
                            <m:r><m:t>−</m:t></m:r>
                            <m:r><m:t>b</m:t></m:r>
                            <m:r><m:t xml:space="preserve"> ± </m:t></m:r>
                            <m:rad>
                                <m:radPr>
                                    <m:deg/>
                                </m:radPr>
                                <m:e>
                                    <m:r><m:t>b</m:t></m:r>
                                    <m:sup><m:r><m:t>2</m:t></m:r></m:sup>
                                    <m:r><m:t xml:space="preserve"> − 4ac</m:t></m:r>
                                </m:e>
                            </m:rad>
                        </m:num>
                    </m:f>
                </m:oMath>
            </m:eq>
        </m:eqArr>
    </m:oMath>
    """
    
    try:
        # Create an instance of the converter
        converter = OmmlConverter(XSL_STYLESHEET_PATH)

        # Perform the conversion
        latex_output = converter.convert_omml_to_latex(sample_omml)

        print("Original OMML XML:")
        print(sample_omml)
        print("-" * 50)
        print("Converted LaTeX:")
        print(latex_output)

    except FileNotFoundError as e:
        print(e)
        print("Please make sure the OMML2MML.XSL path is correct.")

