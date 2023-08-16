import socket
import cv2
import pickle
import struct
import tkinter as tk
from tkinter import ttk
from tkinter import PhotoImage
from PIL import Image, ImageTk


class VideoServerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Video Server App")

        self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.host_ip = '0.0.0.0'
        self.port = 9999
        self.server_socket.bind((self.host_ip, self.port))
        self.server_socket.listen(5)

        self.client_socket = None
        self.data = b""
        self.payload_size = struct.calcsize(">L")

        self.label = ttk.Label(self.root, text="Server đang lắng nghe...")
        self.label.pack(pady=10)

        self.canvas = tk.Canvas(self.root, width=640, height=480)
        self.canvas.pack()

        self.receive_frames()

    def receive_frames(self):
        self.client_socket, client_addr = self.server_socket.accept()
        self.label.config(text=f"Kết nối từ: {client_addr}")

        while True:
            while len(self.data) < self.payload_size:
                self.data += self.client_socket.recv(4096)

            packed_msg_size = self.data[:self.payload_size]
            self.data = self.data[self.payload_size:]
            msg_size = struct.unpack(">L", packed_msg_size)[0]

            while len(self.data) < msg_size:
                self.data += self.client_socket.recv(4096)

            frame_data = self.data[:msg_size]
            self.data = self.data[msg_size:]

            frame = pickle.loads(frame_data)
            self.display_frame(frame)
            self.root.update_idletasks()

    def display_frame(self, frame):
        image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        image = Image.fromarray(image)
        photo = ImageTk.PhotoImage(image=image)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        self.canvas.image = photo


def main():
    root = tk.Tk()
    app = VideoServerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
