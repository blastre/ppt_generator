import argparse
import os
from datetime import datetime
from csv_processor import load_csv, analyze_data
from ppt_generator import create_ppt_skeleton, create_presentation, get_available_templates

# NEW: import chatbot utilities
from chatbot import chatbot, chatbot_loop


def csv_to_ppt(csv_filepath: str, question: str, output_filename: str = None, template: str = "default") -> str:
    """
    Main function to convert CSV and question to PowerPoint presentation

    Args:
        csv_filepath: Path to CSV file
        question: Question/query about the data
        output_filename: Output PPT filename (optional)
        template: Template name to use (default: "default")

    Returns:
        Path to generated PowerPoint file
    """

    print(f"Processing CSV: {csv_filepath}")
    print(f"Question: {question}")
    print(f"Using template: {template}")

    # Validate template
    available_templates = get_available_templates()
    if template not in available_templates:
        print(f"Warning: Template '{template}' not found. Using 'default' template.")
        print(f"Available templates: {', '.join(available_templates)}")
        template = "default"

    # Load CSV data
    try:
        df = load_csv(csv_filepath)
        print(f"Loaded CSV with {df.shape[0]} rows and {df.shape[1]} columns")
    except Exception as e:
        print(f"Error loading CSV: {e}")
        return None

    # Analyze data based on question
    print("Analyzing data...")
    analysis_result = analyze_data(df, question)
    print(f"Analysis complete. Summary: {analysis_result.get('summary', '')}")

    # Create PPT skeleton with template
    print("Creating presentation skeleton...")
    skeleton = create_ppt_skeleton(analysis_result, template)
    print(f"Created skeleton with {len(skeleton.get('slides', []))} slides")

    # Generate output filename if not provided
    if not output_filename:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        template_suffix = f"_{template}" if template != "default" else ""
        output_filename = f"presentation{template_suffix}_{timestamp}.pptx"

    # Create presentation with template
    print("Generating PowerPoint presentation...")
    final_ppt_path = create_presentation(skeleton, analysis_result, output_filename, template)

    print(f"✅ Presentation created successfully: {final_ppt_path}")
    return final_ppt_path


def list_templates():
    """Display available templates"""
    templates = get_available_templates()
    print("\n📋 Available PowerPoint Templates:")
    print("=" * 40)

    template_descriptions = {
        "default": "Default professional template with blue accents",
        "modern_blue": "Modern blue template with clean design",
        "corporate_green": "Corporate green template for business presentations",
        "minimalist_gray": "Minimalist gray template with blue accents",
        "vibrant_orange": "Vibrant orange template for creative presentations",
        "elegant_purple": "Elegant purple template with sophisticated styling"
    }

    for template in templates:
        description = template_descriptions.get(template, "Custom template")
        print(f"  • {template:15} - {description}")

    print("\nUsage: python main.py <csv_file> <question> -t <template_name>")
    print("=" * 40)


def run_chatbot(csv_filepath: str):
    """
    Public function to run chatbot mode given a CSV file.
    This can be imported and used by other programs.

    Args:
        csv_filepath (str): Path to the CSV file
    """
    try:
        df = load_csv(csv_filepath)
        # Create a concise dataset summary for stable context
        analysis_result = analyze_data(
            df, "Please provide a concise factual summary of the dataset."
        )
        summary = analysis_result.get("summary", "")
        # Start interactive chatbot loop
        chatbot_loop(summary)
    except Exception as e:
        print(f"❌ Error starting chatbot mode: {e}")


def main():
    """CLI interface with template support"""
    parser = argparse.ArgumentParser(
        description="Convert CSV data and questions to beautiful PowerPoint presentations",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py data.csv "What are the sales trends?" 
  python main.py data.csv "Analyze customer segments" -t modern_blue
  python main.py data.csv "Revenue analysis" -o my_report.pptx -t corporate_green
  python main.py --list-templates  # Show available templates
  python main.py data.csv --chat   # Start interactive chatbot mode (type /exit to quit)
        """
    )

    parser.add_argument("csv_file", nargs='?', help="Path to CSV file")
    parser.add_argument("question", nargs='?', help="Question about the data")
    parser.add_argument("-o", "--output", help="Output PowerPoint filename", default=None)
    parser.add_argument("-t", "--template", 
                       help="PowerPoint template to use", 
                       default="default",
                       choices=get_available_templates())
    parser.add_argument("--list-templates", action="store_true", 
                       help="List all available templates")
    parser.add_argument("--chat", action="store_true",
                       help="Use chatbot mode: interactive Q&A based on the CSV summary (type /exit to quit)")

    args = parser.parse_args()

    # Handle template listing
    if args.list_templates:
        list_templates()
        return

    # Validate required CSV file argument
    if not args.csv_file:
        parser.print_help()
        print("\n❌ Error: csv_file is required (except when using --list-templates)")
        return

    # Validate CSV file exists
    if not os.path.exists(args.csv_file):
        print(f"❌ Error: CSV file '{args.csv_file}' not found")
        return

    # Handle chatbot mode
    if args.chat:
        run_chatbot(args.csv_file)
        return

    # Validate question for PPT mode
    if not args.question:
        parser.print_help()
        print("\n❌ Error: Both csv_file and question are required for PPT generation")
        print("Use --list-templates to see available templates")
        return

    # Process CSV to PPT
    result = csv_to_ppt(args.csv_file, args.question, args.output, args.template)

    if result:
        print(f"\n🎉 Success! PowerPoint presentation saved as: {result}")
        print(f"📊 Template used: {args.template}")
        print("💡 Tip: Try different templates with -t option for varied styles!")
    else:
        print("\n❌ Failed to create presentation")


if __name__ == "__main__":
    main()

