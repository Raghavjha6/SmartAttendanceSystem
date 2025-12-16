import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import cv2
import face_recognition
import os
import numpy as np
from datetime import datetime, date
import mysql.connector
import shutil
from openpyxl import Workbook

# -----------------------------
# Database Connection
# -----------------------------
db = mysql.connector.connect(
    host="localhost",
    user="root",
    password="root",
    database="smart_attendance"
)
cursor = db.cursor()

# -----------------------------
# Admin Login Class
# -----------------------------
class AdminLogin:
    def __init__(self, root):
        self.root = root
        self.root.title("Admin Login - Smart Attendance System")
        self.root.geometry("400x300")
        self.root.resizable(False, False)

        tk.Label(root, text="Smart Attendance System", font=("Arial", 16, "bold")).pack(pady=15)
        tk.Label(root, text="Admin Login", font=("Arial", 13)).pack(pady=5)

        tk.Label(root, text="Username:").pack(pady=5)
        self.username_entry = tk.Entry(root, width=30)
        self.username_entry.pack()

        tk.Label(root, text="Password:").pack(pady=5)
        self.password_entry = tk.Entry(root, width=30, show="*")
        self.password_entry.pack()

        ttk.Button(root, text="Login", command=self.verify_login).pack(pady=15)

    def verify_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        cursor.execute("SELECT * FROM admin WHERE username=%s AND password=%s", (username, password))
        result = cursor.fetchone()

        if result:
            messagebox.showinfo("Login Successful", f"Welcome, {username}!")
            self.root.destroy()
            app_root = tk.Tk()
            AttendanceApp(app_root, username)
            app_root.mainloop()
        else:
            messagebox.showerror("Error", "Invalid username or password")


# -----------------------------
# Main App Class
# -----------------------------
class AttendanceApp:
    def __init__(self, root, admin_name):
        self.root = root
        self.admin_name = admin_name
        self.root.title(f"Smart Attendance System - Logged in as {admin_name}")
        self.root.geometry("1000x800")

        # Video Frame
        self.video_label = tk.Label(self.root)
        self.video_label.pack()

        # Buttons
        ttk.Button(self.root, text="Start Attendance", command=self.start_attendance).pack(pady=5)
        ttk.Button(self.root, text="Stop Attendance", command=self.stop_attendance).pack(pady=5)
        ttk.Button(self.root, text="Manage Students", command=self.manage_students).pack(pady=5)
        ttk.Button(self.root, text="View Attendance", command=self.view_attendance).pack(pady=5)
        ttk.Button(self.root, text="Export Today's Attendance", command=self.export_today_attendance).pack(pady=5)
        ttk.Button(self.root, text="Export Attendance (Custom Range)", command=self.export_range_attendance).pack(pady=5)

        # NEW Logout Button
        ttk.Button(self.root, text="Logout", command=self.logout).pack(pady=20)

        self.cap = None
        self.running = False

    # -----------------------------
    # Logout Function
    # -----------------------------
    def logout(self):
        self.stop_attendance()
        messagebox.showinfo("Logout", f"Goodbye, {self.admin_name}!")
        self.root.destroy()
        new_root = tk.Tk()
        AdminLogin(new_root)
        new_root.mainloop()

    # -----------------------------
    # Load known faces
    # -----------------------------
    def load_known_faces(self):
        self.known_faces = []
        self.known_names = []
        cursor.execute("SELECT name, image_path FROM students")
        for name, path in cursor.fetchall():
            if os.path.exists(path):
                image = face_recognition.load_image_file(path)
                encoding = face_recognition.face_encodings(image)
                if encoding:
                    self.known_faces.append(encoding[0])
                    self.known_names.append(name)

    # -----------------------------
    # Start Attendance
    # -----------------------------
    def start_attendance(self):
        self.load_known_faces()
        self.cap = cv2.VideoCapture(0)
        self.running = True
        self.video_stream()

    def video_stream(self):
        if not self.running:
            return
        ret, frame = self.cap.read()
        if not ret:
            return

        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
        rgb_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2RGB)

        face_locations = face_recognition.face_locations(rgb_frame)
        face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

        for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
            matches = face_recognition.compare_faces(self.known_faces, face_encoding)
            name = "Unknown"

            if True in matches:
                first_match_index = matches.index(True)
                name = self.known_names[first_match_index]
                self.mark_attendance(name)
            else:
                name = "Not Found"

            top *= 4
            right *= 4
            bottom *= 4
            left *= 4

            color = (0, 255, 0) if name != "Not Found" else (0, 0, 255)
            cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
            cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.9, color, 2)

        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = cv2.resize(cv2image, (700, 500))
        img = tk.PhotoImage(master=self.root, data=cv2.imencode('.png', img)[1].tobytes())
        self.video_label.configure(image=img)
        self.video_label.image = img

        self.root.after(10, self.video_stream)

    # -----------------------------
    # Stop Attendance
    # -----------------------------
    def stop_attendance(self):
        self.running = False
        if self.cap:
            self.cap.release()
        cv2.destroyAllWindows()

    # -----------------------------
    # Mark Attendance (first + last time)
    # -----------------------------
    def mark_attendance(self, student_name):
        cursor.execute("SELECT id FROM students WHERE name=%s", (student_name,))
        result = cursor.fetchone()
        if not result:
            return
        student_id = result[0]
        today = date.today()
        now_time = datetime.now().time()

        cursor.execute("SELECT * FROM attendance WHERE student_id=%s AND date=%s", (student_id, today))
        record = cursor.fetchone()

        if record:
            cursor.execute("UPDATE attendance SET last_time=%s WHERE student_id=%s AND date=%s",
                           (now_time, student_id, today))
        else:
            cursor.execute("INSERT INTO attendance (student_id, date, first_time, last_time) VALUES (%s, %s, %s, %s)",
                           (student_id, today, now_time, now_time))
        db.commit()

    # -----------------------------
    # View Attendance
    # -----------------------------
    def view_attendance(self):
        top = tk.Toplevel(self.root)
        top.title("Attendance Records")
        top.geometry("700x500")

        tree = ttk.Treeview(top, columns=("ID", "Name", "Date", "First Time", "Last Time"), show="headings")
        for col in ("ID", "Name", "Date", "First Time", "Last Time"):
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True)

        cursor.execute("""
            SELECT a.id, s.name, a.date, a.first_time, a.last_time
            FROM attendance a
            JOIN students s ON a.student_id = s.id
            ORDER BY a.date DESC, s.name
        """)
        for row in cursor.fetchall():
            tree.insert("", "end", values=row)

    # -----------------------------
    # Export Functions
    # -----------------------------
    def export_today_attendance(self):
        today = date.today()
        cursor.execute("""
            SELECT s.name, s.roll_no, a.date, a.first_time, a.last_time
            FROM attendance a
            JOIN students s ON a.student_id = s.id
            WHERE a.date = %s
        """, (today,))
        data = cursor.fetchall()
        if not data:
            messagebox.showinfo("No Data", "No attendance data found for today.")
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Today's Attendance"
        ws.append(["Name", "Roll No", "Date", "First Time", "Last Time"])
        for row in data:
            ws.append(row)

        filename = f"attendance_{today}.xlsx"
        wb.save(filename)
        messagebox.showinfo("Success", f"Attendance exported as {filename}")

    def export_range_attendance(self):
        top = tk.Toplevel(self.root)
        top.title("Export Attendance Range")

        tk.Label(top, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, padx=10, pady=5)
        start_entry = tk.Entry(top)
        start_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Label(top, text="End Date (YYYY-MM-DD):").grid(row=1, column=0, padx=10, pady=5)
        end_entry = tk.Entry(top)
        end_entry.grid(row=1, column=1, padx=10, pady=5)

        def export():
            start_date = start_entry.get()
            end_date = end_entry.get()
            cursor.execute("""
                SELECT s.name, s.roll_no, a.date, a.first_time, a.last_time
                FROM attendance a
                JOIN students s ON a.student_id = s.id
                WHERE a.date BETWEEN %s AND %s
                ORDER BY a.date, s.name
            """, (start_date, end_date))
            data = cursor.fetchall()
            if not data:
                messagebox.showinfo("No Data", "No records found in this range.")
                return

            wb = Workbook()
            ws = wb.active
            ws.title = "Attendance Range"
            ws.append(["Name", "Roll No", "Date", "First Time", "Last Time"])
            for row in data:
                ws.append(row)

            filename = f"attendance_{start_date}_to_{end_date}.xlsx"
            wb.save(filename)
            messagebox.showinfo("Success", f"Attendance exported as {filename}")
            top.destroy()

        ttk.Button(top, text="Export", command=export).grid(row=2, column=0, columnspan=2, pady=10)

    # -----------------------------
    # Manage Students
    # -----------------------------
    def manage_students(self):
        top = tk.Toplevel(self.root)
        top.title("Student Management")
        top.geometry("700x500")

        tree = ttk.Treeview(top, columns=("ID", "Name", "Roll No", "Image"), show="headings")
        for col in ("ID", "Name", "Roll No", "Image"):
            tree.heading(col, text=col)
        tree.pack(fill="both", expand=True, pady=10)

        def load_students():
            tree.delete(*tree.get_children())
            cursor.execute("SELECT * FROM students")
            for row in cursor.fetchall():
                tree.insert("", "end", values=row)

        load_students()

        def add_student():
            name = name_entry.get()
            roll_no = roll_entry.get()
            image_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.jpg;*.png")])
            if not name or not roll_no or not image_path:
                messagebox.showerror("Error", "All fields required")
                return
            dest_path = f"known_faces/{name}.jpg"
            shutil.copy(image_path, dest_path)
            cursor.execute("INSERT INTO students (name, roll_no, image_path) VALUES (%s, %s, %s)",
                           (name, roll_no, dest_path))
            db.commit()
            messagebox.showinfo("Success", "Student Added Successfully")
            load_students()

        def update_student():
            selected = tree.focus()
            if not selected:
                messagebox.showerror("Error", "Select a student first")
                return
            values = tree.item(selected, "values")
            student_id = values[0]
            new_name = name_entry.get()
            new_roll = roll_entry.get()
            cursor.execute("UPDATE students SET name=%s, roll_no=%s WHERE id=%s",
                           (new_name, new_roll, student_id))
            db.commit()
            messagebox.showinfo("Success", "Student Updated")
            load_students()

        def delete_student():
            selected = tree.focus()
            if not selected:
                messagebox.showerror("Error", "Select a student first")
                return
            values = tree.item(selected, "values")
            student_id = values[0]
            cursor.execute("DELETE FROM students WHERE id=%s", (student_id,))
            db.commit()
            messagebox.showinfo("Deleted", "Student Deleted")
            load_students()

        form_frame = tk.Frame(top)
        form_frame.pack(pady=10)
        tk.Label(form_frame, text="Name:").grid(row=0, column=0, padx=5)
        name_entry = tk.Entry(form_frame)
        name_entry.grid(row=0, column=1, padx=5)
        tk.Label(form_frame, text="Roll No:").grid(row=0, column=2, padx=5)
        roll_entry = tk.Entry(form_frame)
        roll_entry.grid(row=0, column=3, padx=5)
        ttk.Button(form_frame, text="Add Student", command=add_student).grid(row=1, column=0, pady=10)
        ttk.Button(form_frame, text="Update Student", command=update_student).grid(row=1, column=1, pady=10)
        ttk.Button(form_frame, text="Delete Student", command=delete_student).grid(row=1, column=2, pady=10)


# -----------------------------
# Run the App
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    AdminLogin(root)
    root.mainloop()
