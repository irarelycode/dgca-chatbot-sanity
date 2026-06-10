"""
Run both Dashboard PDF and Chatbot Sanity, then email both reports.
"""

from utils import send_email, OUTPUT_DIR
import dashboard_pdf
import chatbot_sanity

def main():
    print("=" * 80)
    print("DGCA COMPLETE AUTOMATION – Dashboard + Chatbot Sanity")
    print("=" * 80)

    attachments = []

    # 1. Dashboard PDF
    try:
        pdf_path = dashboard_pdf.main()
        attachments.append(pdf_path)
        print(f"\nDashboard PDF: {pdf_path}")
    except Exception as e:
        print(f"Dashboard PDF failed: {e}")

    # 2. Chatbot Sanity
    try:
        sanity_path = chatbot_sanity.main()
        attachments.append(sanity_path)
        print(f"Sanity report: {sanity_path}")
    except Exception as e:
        print(f"Chatbot sanity failed: {e}")

    # 3. Email
    if attachments:
        send_email(attachments, subject="DGCA Daily Automation Report")
        print("\nEmail sent with both reports.")
    else:
        print("\nNo reports generated – email not sent.")

if __name__ == "__main__":
    main()
