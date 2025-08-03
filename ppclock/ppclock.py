#!/usr/bin/env python3
"""
PPClock - Countdown Timer for PowerPoint
Phase 1: Desktop Application

Features:
- Input dialog for setting countdown time
- Movable popup countdown display
- Clean, professional appearance
"""

import tkinter as tk
from tkinter import simpledialog, messagebox
import threading
import time


class CountdownTimer:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()  # Hide main window
        self.countdown_window = None
        self.is_running = False
        self.remaining_seconds = 0
        
    def get_time_input(self):
        """Get countdown time from user via input dialog"""
        try:
            # Ask for minutes
            minutes = simpledialog.askinteger(
                "PPClock Setup",
                "Enter countdown time in minutes:",
                minvalue=1,
                maxvalue=999
            )
            
            if minutes is None:  # User cancelled
                return None
                
            # Ask for additional seconds
            seconds = simpledialog.askinteger(
                "PPClock Setup", 
                "Enter additional seconds (0-59):",
                minvalue=0,
                maxvalue=59,
                initialvalue=0
            )
            
            if seconds is None:  # User cancelled
                return None
                
            total_seconds = (minutes * 60) + seconds
            return total_seconds
            
        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")
            return None
    
    def format_time(self, seconds):
        """Format seconds into MM:SS or HH:MM:SS"""
        hours = seconds // 3600
        minutes = (seconds % 3600) // 60
        secs = seconds % 60
        
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{secs:02d}"
        else:
            return f"{minutes:02d}:{secs:02d}"
    
    def create_countdown_window(self):
        """Create the movable countdown display window"""
        self.countdown_window = tk.Toplevel(self.root)
        self.countdown_window.title("PPClock")
        self.countdown_window.geometry("200x80+100+100")  # width x height + x_offset + y_offset
        
        # Make window stay on top but allow moving
        self.countdown_window.attributes("-topmost", True)
        
        # Configure window styling
        self.countdown_window.configure(bg='#2c3e50')
        
        # Create time display label
        self.time_label = tk.Label(
            self.countdown_window,
            text=self.format_time(self.remaining_seconds),
            font=('Arial', 20, 'bold'),
            fg='#ecf0f1',
            bg='#2c3e50'
        )
        self.time_label.pack(expand=True)
        
        # Add control buttons
        button_frame = tk.Frame(self.countdown_window, bg='#2c3e50')
        button_frame.pack(side='bottom', fill='x', padx=5, pady=5)
        
        self.pause_button = tk.Button(
            button_frame,
            text="Pause",
            command=self.toggle_pause,
            bg='#f39c12',
            fg='white',
            relief='flat'
        )
        self.pause_button.pack(side='left', padx=2)
        
        self.stop_button = tk.Button(
            button_frame,
            text="Stop",
            command=self.stop_countdown,
            bg='#e74c3c',
            fg='white',
            relief='flat'
        )
        self.stop_button.pack(side='right', padx=2)
        
        # Handle window close
        self.countdown_window.protocol("WM_DELETE_WINDOW", self.stop_countdown)
    
    def countdown_worker(self):
        """Background thread that handles the countdown logic"""
        while self.remaining_seconds > 0 and self.is_running:
            if not hasattr(self, 'paused') or not self.paused:
                # Update display
                if self.countdown_window and self.time_label:
                    self.time_label.config(text=self.format_time(self.remaining_seconds))
                
                time.sleep(1)
                self.remaining_seconds -= 1
            else:
                time.sleep(0.1)  # Short sleep when paused
        
        # Countdown finished
        if self.remaining_seconds <= 0 and self.is_running:
            self.countdown_finished()
    
    def countdown_finished(self):
        """Handle countdown completion"""
        if self.countdown_window and self.time_label:
            self.time_label.config(text="00:00", fg='#e74c3c')
        
        # Show completion message
        messagebox.showinfo("PPClock", "Time's up!")
        self.stop_countdown()
    
    def toggle_pause(self):
        """Toggle pause/resume"""
        if not hasattr(self, 'paused'):
            self.paused = False
            
        self.paused = not self.paused
        
        if self.paused:
            self.pause_button.config(text="Resume", bg='#27ae60')
        else:
            self.pause_button.config(text="Pause", bg='#f39c12')
    
    def stop_countdown(self):
        """Stop the countdown and close window"""
        self.is_running = False
        if self.countdown_window:
            self.countdown_window.destroy()
            self.countdown_window = None
        self.root.quit()
    
    def start_countdown(self):
        """Main method to start the countdown process"""
        # Get time input from user
        total_seconds = self.get_time_input()
        
        if total_seconds is None:
            self.root.quit()
            return
        
        self.remaining_seconds = total_seconds
        self.is_running = True
        self.paused = False
        
        # Create countdown window
        self.create_countdown_window()
        
        # Start countdown in background thread
        countdown_thread = threading.Thread(target=self.countdown_worker, daemon=True)
        countdown_thread.start()
        
        # Start GUI main loop
        self.root.mainloop()


def main():
    """Entry point for PPClock application"""
    print("Starting PPClock - Countdown Timer")
    timer = CountdownTimer()
    timer.start_countdown()


if __name__ == "__main__":
    main()