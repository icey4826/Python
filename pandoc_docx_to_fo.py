import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import argparse
import shutil
import jpype
import jpype.imports
from pathlib import Path

def select_files():
    """Show file selection dialogs and return input/output paths"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Ensure dialogs appear on top
    
    # Select input file
    input_file = filedialog.askopenfilename(
        title="Select Word Document",
        filetypes=[
            ("Word documents", "*.docx"),
            ("All files", "*.*")
        ]
    )
    
    if not input_file:
        return None, None
        
    # Select output file
    output_file = filedialog.asksaveasfilename(
        title="Save XSL-FO file as",
        defaultextension=".fo",
        filetypes=[
            ("XSL-FO files", "*.fo"),
            ("All files", "*.*")
        ],
        initialfile="output.fo"
    )
    
    root.destroy()  # Clean up the tk instance
    return input_file, output_file

class Docx4jConverter:
    def __init__(self):
        self.verify_java()
        self.init_java()
        
    def verify_java(self):
        """Verify that Java is installed and accessible"""
        try:
            result = subprocess.run(['java', '-version'], 
                                  capture_output=True, 
                                  text=True)
            if result.returncode != 0:
                raise Exception("Java is not installed or not in PATH")
        except FileNotFoundError:
            raise Exception(
                "Java is not installed or not in PATH. "
                "Please install Java from https://adoptium.net/"
            )

    def init_java(self):
        """Initialize Java and required docx4j libraries"""
        if not jpype.isJVMStarted():
            # Download docx4j JARs if not present
            self.ensure_jars_exist()
            
            # Start JVM with docx4j
            jars_dir = os.path.join(os.path.dirname(__file__), 'lib')
            classpath = os.pathsep.join([str(x) for x in Path(jars_dir).glob("*.jar")])
            jpype.startJVM(classpath=classpath, convertStrings=True)

    def ensure_jars_exist(self):
        """Download required docx4j JARs if they don't exist"""
        lib_dir = os.path.join(os.path.dirname(__file__), 'lib')
        os.makedirs(lib_dir, exist_ok=True)
        
        # Path to local Maven installation
        maven_path = os.path.join(os.path.dirname(__file__), 'apache-maven-3.9.9', 'bin', 'mvn')
        if os.name == 'nt':  # Windows
            maven_path += '.cmd'
        
        # List of required JARs and their Maven coordinates
        required_jars = [
            # Core docx4j dependencies
            "org.docx4j:docx4j-core:11.4.9",
            "org.docx4j:docx4j-export-fo:11.4.9",
            "org.docx4j:docx4j-JAXB-Internal:11.4.9",
            "org.docx4j:docx4j-JAXB-ReferenceImpl:11.4.9",
            "org.docx4j:docx4j-JAXB-MOXy:11.4.9",
            
            # JAXB dependencies
            "jakarta.xml.bind:jakarta.xml.bind-api:4.0.0",
            "org.glassfish.jaxb:jaxb-runtime:4.0.3",
            "org.glassfish.jaxb:jaxb-core:4.0.3",
            "com.sun.xml.bind:jaxb-impl:4.0.3",
            
            # Additional required dependencies
            "org.slf4j:slf4j-api:2.0.9",
            "org.slf4j:slf4j-simple:2.0.9",
            "commons-io:commons-io:2.11.0",
            "org.apache.commons:commons-lang3:3.12.0",
            "org.apache.xmlgraphics:fop:2.8",
            "org.apache.xmlgraphics:xmlgraphics-commons:2.8",
            "xalan:xalan:2.7.2",
            "xerces:xercesImpl:2.12.2"
        ]
        
        # Use Maven to download JARs
        for jar in required_jars:
            jar_path = os.path.join(lib_dir, f"{jar.split(':')[1]}-{jar.split(':')[2]}.jar")
            if not os.path.exists(jar_path):
                try:
                    subprocess.run([
                        maven_path,
                        "org.apache.maven.plugins:maven-dependency-plugin:2.10:get",
                        f"-Dartifact={jar}",
                        f"-Ddest={jar_path}"
                    ], check=True)
                except subprocess.CalledProcessError as e:
                    raise Exception(f"Failed to download JAR {jar}: {str(e)}")
                except FileNotFoundError:
                    raise Exception(
                        f"Maven not found at {maven_path}. Please ensure Maven is extracted to "
                        "the correct location relative to this script."
                    )

    def convert(self, input_file, output_file):
        """Convert Word document to XSL-FO using docx4j"""
        try:
            # Import required Java classes
            from org.docx4j.openpackaging.packages import WordprocessingMLPackage
            from org.docx4j.convert.out.fo import FOSettings
            
            # Load the DOCX
            wordMLPackage = WordprocessingMLPackage.load(jpype.JFile(input_file))
            
            # Configure FO settings
            foSettings = FOSettings()
            
            # Convert to XSL-FO
            with open(output_file, 'wb') as fo_file:
                fo_file.write(
                    wordMLPackage.convert(foSettings, jpype.JClass('org.docx4j.convert.out.fo.FOExporterXslt').class_)
                )
                
        except Exception as e:
            raise Exception(f"Conversion failed: {str(e)}")

def main():
    # Check if command line arguments are provided
    if len(sys.argv) > 1:
        # Use command line arguments
        parser = argparse.ArgumentParser(description='Convert Word DOCX to XSL-FO format using docx4j')
        parser.add_argument('input', help='Input DOCX file path')
        parser.add_argument('output', help='Output XSL-FO file path')
        args = parser.parse_args()
        input_file = args.input
        output_file = args.output
    else:
        # Use file selection dialog
        input_file, output_file = select_files()
        if not input_file or not output_file:
            print("File selection cancelled", file=sys.stderr)
            sys.exit(1)
    
    # Validate input file
    if not os.path.exists(input_file):
        error_msg = f"Error: Input file '{input_file}' does not exist"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)
        
    # Validate output directory exists
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
        except Exception as e:
            error_msg = f"Error creating output directory: {str(e)}"
            if len(sys.argv) > 1:
                print(error_msg, file=sys.stderr)
            else:
                messagebox.showerror("Error", error_msg)
            sys.exit(1)
            
    # Perform conversion
    try:
        converter = Docx4jConverter()
        converter.convert(input_file, output_file)
        success_msg = f"Successfully converted '{input_file}' to '{output_file}'"
        if len(sys.argv) > 1:
            print(success_msg)
        else:
            messagebox.showinfo("Success", success_msg)
    except Exception as e:
        error_msg = f"Error during conversion: {str(e)}"
        if len(sys.argv) > 1:
            print(error_msg, file=sys.stderr)
        else:
            messagebox.showerror("Error", error_msg)
        sys.exit(1)

if __name__ == '__main__':
    main() 