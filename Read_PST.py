#!/usr/bin/env python3

import argparse
import csv
import uuid
from pathlib import Path
import pypff  # pip install libpff-python
from datetime import datetime
from tqdm import tqdm


def extract_messages(pst_path: Path, output_dir: Path):
    pst_file = pypff.file()
    pst_file.open(str(pst_path))

    root_folder = pst_file.get_root_folder()
    messages = []

    def walk_folder(folder):
        num_messages = folder.number_of_sub_messages
        for i in tqdm(range(num_messages), desc="Extracting messages", unit="msg", leave=False):
            try:
                message = folder.get_sub_message(i)
                email_id = str(uuid.uuid4())

                try:
                    subject = message.subject or ""
                except OSError:
                    subject = "[Unreadable Subject]"

                try:
                    sender = message.sender_name or ""
                except OSError:
                    sender = "[Unreadable Sender]"

                try:
                    date_obj = message.delivery_time or message.creation_time
                    date_received = date_obj.isoformat() if date_obj else ""
                except OSError:
                    date_received = "[Unreadable Date]"

                try:
                    full_body_html = message.html_body
                    if isinstance(full_body_html, bytes):
                        full_body_html = full_body_html.decode(errors="replace")
                except Exception:
                    full_body_html = None

                if not full_body_html:
                    try:
                        full_body_html = str(message.plain_text_body or "")
                        if isinstance(full_body_html, bytes):
                            full_body_html = full_body_html.decode(errors="replace")
                        full_body_html = "<pre>" + full_body_html + "</pre>"
                    except Exception:
                        full_body_html = "<p>[Unreadable Body]</p>"

                # Save HTML body to folder
                email_folder = output_dir / email_id
                email_folder.mkdir(parents=True, exist_ok=True)
                with open(email_folder / "body.html", "w", encoding="utf-8") as f:
                    f.write(full_body_html)

                # Save truncated plain body for CSV
                try:
                    truncated_body = str(message.plain_text_body or "")
                    if isinstance(truncated_body, bytes):
                        truncated_body = truncated_body.decode(errors="replace")
                    truncated_body = truncated_body.strip().replace("\n", " ")[:500]
                except Exception:
                    truncated_body = "[Unreadable Body]"

                messages.append({
                    "email_id": email_id,
                    "date_received": date_received,
                    "subject": subject,
                    "sender": sender,
                    "body": truncated_body
                })

            except Exception as e:
                print(f"Error reading message {i}: {e}")

        for i in range(folder.number_of_sub_folders):
            walk_folder(folder.get_sub_folder(i))


    walk_folder(root_folder)

    # Sort messages by date_received (newest first)
    def sort_key(msg):
        try:
            return datetime.fromisoformat(msg["date_received"])
        except Exception:
            return datetime.min

    messages.sort(key=sort_key, reverse=True)
    return messages


def write_csv(messages, csv_path: Path):
    with open(csv_path, mode="w", newline='', encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["email_id", "date_received", "subject", "sender", "body"])
        writer.writeheader()
        writer.writerows(messages)


def main():
    parser = argparse.ArgumentParser(
        description="Extract emails from a .pst file, export metadata to a CSV, and save each full email body to its own folder."
    )
    parser.add_argument("-pst_path", type=Path, required=True, help="Path to the input .pst file")
    parser.add_argument("-output_path", type=Path, required=True, help="Path to the output folder where CSV and email folders will be saved")

    args = parser.parse_args()

    if not args.pst_path.exists():
        parser.error(f"The file '{args.pst_path}' does not exist.")
    if args.output_path.exists():
        parser.error(f"The output directory '{args.output_path}' already exists. Please choose a new directory.")
    args.output_path.mkdir(parents=True, exist_ok=True)

    # Derive CSV file name from PST file
    csv_name = args.pst_path.stem + ".csv"
    csv_path = args.output_path / csv_name

    messages = extract_messages(args.pst_path, args.output_path)
    write_csv(messages, csv_path)

    print(f"Exported {len(messages)} messages.")
    print(f"CSV saved to: {csv_path}")
    print(f"Email folders created in: {args.output_path}")


if __name__ == "__main__":
    main()
