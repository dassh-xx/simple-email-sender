import smtplib
import ssl
import os
import time
from email.message import EmailMessage
from email.utils import formataddr
from mimetypes import guess_type
import tkinter as tk
from tkinter import filedialog, messagebox, StringVar, IntVar, Text
from tkinter import ttk

attachments = []
smtp_entries = []

# Build email
def build_email(subject, sender_name, sender_email, reply_to, body, is_html, attachments, recipient):
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = formataddr((sender_name, sender_email))
    msg['Reply-To'] = reply_to
    msg['To'] = recipient

    if is_html:
        msg.set_content("This is the plain text fallback version.", subtype='plain')
        msg.add_alternative(body, subtype='html')
    else:
        msg.set_content(body, subtype='plain')

    for filepath in attachments:
        with open(filepath, 'rb') as f:
            data = f.read()
            mime_type = guess_type(filepath)[0]
            if mime_type:
                maintype, subtype = mime_type.split('/')
            else:
                maintype, subtype = 'application', 'octet-stream'
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(filepath))
    return msg

# Send email
def send_email(smtp_info, message):
    try:
        context = ssl.create_default_context()
        if smtp_info['port'] == 465:
            with smtplib.SMTP_SSL(smtp_info['server'], smtp_info['port'], context=context) as server:
                server.login(smtp_info['username'], smtp_info['password'])
                server.send_message(message)
        elif smtp_info['port'] == 587:
            with smtplib.SMTP(smtp_info['server'], smtp_info['port']) as server:
                server.starttls(context=context)
                server.login(smtp_info['username'], smtp_info['password'])
                server.send_message(message)
        return True
    except Exception as e:
        log(f"❌ Failed to send to {message['To']}: {e}")
        return False

def log(message):
    output.insert(tk.END, message + "\n")
    output.see(tk.END)

def start_sending():
    subject = subject_entry.get().strip()
    sender_name = sender_entry.get().strip()
    reply_to = replyto_entry.get().strip()
    recipients_file = recipients_path.get()
    body_file = body_path.get()
    is_html = body_type.get()
    delay = int(delay_entry.get().strip() or "0")
    use_rotation = rotation_var.get()

    smtp_list = []
    if use_rotation:
        for row in smtp_entries:
            server = row['server'].get().strip()
            port = row['port'].get().strip()
            user = row['user'].get().strip()
            pwd = row['pass'].get().strip()
            if not all([server, port, user, pwd]):
                messagebox.showerror("Error", "All SMTP fields must be filled in rotation mode.")
                return
            smtp_list.append({
                "server": server,
                "port": int(port),
                "username": user,
                "password": pwd
            })
    else:
        smtp_list.append({
            "server": smtp_server_entry.get().strip(),
            "port": int(smtp_port_entry.get().strip()),
            "username": smtp_user_entry.get().strip(),
            "password": smtp_pass_entry.get().strip()
        })

    if not smtp_list:
        messagebox.showerror("Error", "No SMTP accounts configured.")
        return

    if not all([subject, sender_name, reply_to, recipients_file, body_file]):
        messagebox.showerror("Error", "Please fill all fields and select files.")
        return

    with open(recipients_file, "r", encoding="utf-8") as f:
        recipients = [line.strip() for line in f if line.strip()]
    with open(body_file, "r", encoding="utf-8") as f:
        body = f.read()

    sent = []
    failed = []

    log("🚀 Starting email sending...")

    smtp_index = 0

    for recipient in recipients:
        smtp_info = smtp_list[smtp_index]
        smtp_index = (smtp_index + 1) % len(smtp_list)

        msg = build_email(
            subject,
            sender_name,
            smtp_info["username"],
            reply_to,
            body,
            is_html,
            attachments,
            recipient
        )
        success = send_email(smtp_info, msg)
        if success:
            log(f"✅ Sent to {recipient}")
            sent.append(recipient)
        else:
            log(f"❌ Failed to {recipient}")
            failed.append(recipient)
        root.update()
        time.sleep(delay)

    with open("sent.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(sent))
    with open("failed.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(failed))

    messagebox.showinfo("Done", f"Sent: {len(sent)}\nFailed: {len(failed)}\nSee sent.txt and failed.txt.")

def select_file(var):
    path = filedialog.askopenfilename()
    if path:
        var.set(path)

def add_attachment():
    path = filedialog.askopenfilename()
    if path:
        attachments.append(path)
        log(f"📎 Added attachment: {os.path.basename(path)}")

def toggle_rotation():
    if rotation_var.get():
        smtp_frame.grid()
        single_smtp_frame.grid_remove()
    else:
        smtp_frame.grid_remove()
        single_smtp_frame.grid()

def add_smtp_row():
    row_frame = tk.Frame(smtp_frame, bg="#1f2937")
    row_frame.pack(pady=2, fill="x")

    server = tk.Entry(row_frame, **entry_style, width=15)
    server.pack(side="left")
    port = tk.Entry(row_frame, **entry_style, width=4)
    port.insert(0, "587")
    port.pack(side="left", padx=2)
    user = tk.Entry(row_frame, **entry_style, width=20)
    user.pack(side="left", padx=2)
    pwd = tk.Entry(row_frame, **entry_style, width=15, show="*")
    pwd.pack(side="left", padx=2)

    btn_remove = ttk.Button(row_frame, text="X", command=lambda: remove_smtp_row(row_frame))
    btn_remove.pack(side="left", padx=2)

    smtp_entries.append({"frame": row_frame, "server": server, "port": port, "user": user, "pass": pwd})

def remove_smtp_row(frame):
    for i, entry in enumerate(smtp_entries):
        if entry['frame'] == frame:
            smtp_entries.pop(i)
            break
    frame.destroy()

root = tk.Tk()
root.title("dash simple email sender")
root.configure(bg="#1f2937")
root.geometry("900x700")

FONT = ("Segoe UI", 10)

# Custom style
style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", background="#1f2937", foreground="#f3f4f6", font=FONT)
style.configure("TButton", background="#1f2937", foreground="#f9fafb", font=FONT, padding=6)
style.map("TButton", background=[("active", "#1f2937")])
style.configure("TCheckbutton", background="#1f2937", foreground="#f3f4f6", font=FONT)

entry_style = {
    "bg": "#374151",
    "fg": "#f9fafb",
    "insertbackground": "#f9fafb",
    "relief": "flat",
    "highlightthickness": 0,
    "font": FONT
}

# Subject
ttk.Label(root, text="Subject:").grid(row=0, column=0, sticky="w", padx=12, pady=6)
subject_entry = tk.Entry(root, **entry_style, width=60)
subject_entry.grid(row=0, column=1, columnspan=3, padx=12, pady=6)

# Sender and reply
ttk.Label(root, text="Sender Name:").grid(row=1, column=0, sticky="w", padx=12)
sender_entry = tk.Entry(root, **entry_style)
sender_entry.grid(row=1, column=1, padx=12)

ttk.Label(root, text="Reply-To:").grid(row=1, column=2, sticky="w", padx=12)
replyto_entry = tk.Entry(root, **entry_style)
replyto_entry.grid(row=1, column=3, padx=12)

# Body
body_type = IntVar(value=0)
ttk.Checkbutton(root, text="HTML Body", variable=body_type).grid(row=2, column=0, sticky="w", padx=12, pady=4)

body_path = StringVar()
ttk.Button(root, text="Select Body File", command=lambda: select_file(body_path)).grid(row=2, column=1, padx=4)
ttk.Label(root, textvariable=body_path).grid(row=2, column=2, columnspan=2, sticky="w", padx=4)

recipients_path = StringVar()
ttk.Button(root, text="Select Recipients File", command=lambda: select_file(recipients_path)).grid(row=3, column=1, padx=4)
ttk.Label(root, textvariable=recipients_path).grid(row=3, column=2, columnspan=2, sticky="w", padx=4)

ttk.Button(root, text="Add Attachment", command=add_attachment).grid(row=4, column=0, sticky="w", padx=12, pady=4)

ttk.Label(root, text="Delay (s):").grid(row=4, column=1, sticky="e")
delay_entry = tk.Entry(root, **entry_style, width=5)
delay_entry.insert(0, "1")
delay_entry.grid(row=4, column=2, sticky="w", padx=4)

rotation_var = IntVar(value=0)
ttk.Checkbutton(root, text="Enable SMTP Rotation", variable=rotation_var, command=toggle_rotation).grid(row=5, column=0, sticky="w", padx=12, pady=4)

# Multi SMTP
smtp_frame = tk.Frame(root, bg="#1f2937")
smtp_frame.grid(row=6, column=0, columnspan=4, sticky="we", padx=8, pady=6)
header = tk.Frame(smtp_frame, bg="#1f2937")
header.pack(fill="x", pady=2)
ttk.Label(header, text="Server", width=15).pack(side="left")
ttk.Label(header, text="Port", width=4).pack(side="left", padx=2)
ttk.Label(header, text="Username", width=20).pack(side="left", padx=2)
ttk.Label(header, text="Password", width=15).pack(side="left", padx=2)

ttk.Button(smtp_frame, text="Add SMTP Account", command=add_smtp_row).pack(anchor="w", pady=2)
smtp_frame.grid_remove()

# Single SMTP styled like Multi SMTP
single_smtp_frame = tk.Frame(root, bg="#1f2937")
single_smtp_frame.grid(row=7, column=0, columnspan=4, sticky="we", padx=8, pady=6)

# Row 1: Server + Port
row1 = tk.Frame(single_smtp_frame, bg="#1f2937")
row1.pack(fill="x", pady=2)

tk.Label(row1, text="SMTP Server:", fg="#f3f4f6", bg="#1f2937", font=FONT, width=12, anchor="w").pack(side="left")
smtp_server_entry = tk.Entry(row1, **entry_style, width=25)
smtp_server_entry.pack(side="left", padx=4)

tk.Label(row1, text="Port:", fg="#f3f4f6", bg="#1f2937", font=FONT).pack(side="left", padx=(10,2))
smtp_port_entry = tk.Entry(row1, **entry_style, width=5)
smtp_port_entry.insert(0, "587")
smtp_port_entry.pack(side="left", padx=4)

# Row 2: Username + Password
row2 = tk.Frame(single_smtp_frame, bg="#1f2937")
row2.pack(fill="x", pady=2)

tk.Label(row2, text="SMTP Username:", fg="#f3f4f6", bg="#1f2937", font=FONT, width=12, anchor="w").pack(side="left")
smtp_user_entry = tk.Entry(row2, **entry_style, width=25)
smtp_user_entry.pack(side="left", padx=4)

tk.Label(row2, text="SMTP Password:", fg="#f3f4f6", bg="#1f2937", font=FONT).pack(side="left", padx=(10,2))
smtp_pass_entry = tk.Entry(row2, **entry_style, width=20, show="*")
smtp_pass_entry.pack(side="left", padx=4)

# Send button
ttk.Button(root, text="Send Emails", command=start_sending).grid(row=8, column=0, columnspan=4, sticky="we", padx=12, pady=8)

# Log output
output = Text(root, bg="#111827", fg="#f9fafb", font=("Consolas", 10), relief="flat", wrap="word")
output.grid(row=9, column=0, columnspan=4, padx=12, pady=8, sticky="nsew")

root.grid_rowconfigure(9, weight=1)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_columnconfigure(3, weight=1)

root.mainloop()
