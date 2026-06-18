import cv2
from ultralytics import YOLO
import numpy as np
import time
import os
from datetime import datetime

# Global variables for ROI
roi_points = []
drawing = False

# Dictionary to store entry times for IDs in ROI: {track_id: start_time}
roi_entry_times = {}
# Threshold for loitering (seconds)
LOITERING_THRESHOLD = 5.0
# Cooldown for alerts (seconds) to prevent spam
last_alert_time = 0
ALERT_COOLDOWN = 3.0

# Create alerts directory
if not os.path.exists("alerts"):
    os.makedirs("alerts")

def mouse_callback(event, x, y, flags, param):
    global roi_points, drawing
    
    if event == cv2.EVENT_LBUTTONDOWN:
        drawing = True
        roi_points.append((x, y))
    elif event == cv2.EVENT_RBUTTONDOWN:
        # Right click to clear ROI
        roi_points = []
        drawing = False

def is_inside_roi(point, polygon):
    if len(polygon) < 3:
        return False
    return cv2.pointPolygonTest(np.array(polygon), point, False) >= 0

def send_telegram_alert(image_path, message):
    # Placeholder for Telegram API integration
    # You would use 'requests' library and your bot token here
    print(f"--- TELEGRAM SENT: {message} | Image: {image_path} ---")

def main():
    global last_alert_time
    
    # Load the YOLOv8 model
    print("Loading model...")
    model = YOLO('yolov8n.pt') 

    # Open the webcam
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("Error: Could not open camera.")
        return

    # Set resolution
    width = 1280
    height = 720
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)

    # Setup mouse callback
    cv2.namedWindow("Theft Detection System")
    cv2.setMouseCallback("Theft Detection System", mouse_callback)

    print("System ready.")
    print("- Left click: Add restricted area (ROI) point.")
    print("- Right click: Clear restricted area.")
    print("- 'q': Quit.")

    while True:
        start_time = time.time()
        ret, frame = cap.read()
        if not ret:
            break

        # Run YOLOv8 tracking
        results = model.track(frame, persist=True, verbose=False, classes=[0])

        current_ids_in_roi = set()

        # Draw ROI
        if len(roi_points) > 0:
            cv2.polylines(frame, [np.array(roi_points)], isClosed=True, color=(0, 255, 255), thickness=2)
            overlay = frame.copy()
            cv2.fillPoly(overlay, [np.array(roi_points)], (0, 255, 255))
            cv2.addWeighted(overlay, 0.3, frame, 0.7, 0, frame)

        detection_alert = False
        alert_message = ""

        if results[0].boxes.id is not None:
            track_ids = results[0].boxes.id.int().cpu().tolist()
            cls_ids = results[0].boxes.cls.int().cpu().tolist()
            coords = results[0].boxes.xyxy.cpu().tolist()

            for box, track_id, cls_id in zip(coords, track_ids, cls_ids):
                x1, y1, x2, y2 = map(int, box)
                center_x = int((x1 + x2) / 2)
                center_y = int(y2)
                
                inside = is_inside_roi((center_x, center_y), roi_points)
                
                color = (0, 255, 0)
                status = "Normal"
                
                if inside:
                    current_ids_in_roi.add(track_id)
                    # Track entry time
                    if track_id not in roi_entry_times:
                        roi_entry_times[track_id] = time.time()
                    
                    duration = time.time() - roi_entry_times[track_id]
                    
                    # Loitering Check
                    if duration > LOITERING_THRESHOLD:
                        color = (0, 0, 255)
                        status = "SUSPICIOUS (LOITERING)!"
                        detection_alert = True
                        alert_message = f"ID:{track_id} loitering in restricted area for {duration:.1f}s!"
                    else:
                        color = (0, 165, 255) # Orange for warning
                        status = f"Warning ({duration:.1f}s)"
                        
                    cv2.circle(frame, (center_x, center_y), 5, (0, 0, 255), -1)
                else:
                    # Remove from tracking if exited ROI
                    if track_id in roi_entry_times:
                        roi_entry_times.pop(track_id)

                cv2.rectangle(frame, (x1, y1), (x2, y2), color, 2)
                cv2.putText(frame, f"ID:{track_id} {status}", (x1, y1 - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

        # Cleanup roi_entry_times for IDs that are no longer detected at all
        current_track_ids = set(results[0].boxes.id.int().cpu().tolist()) if results[0].boxes.id is not None else set()
        for tid in list(roi_entry_times.keys()):
            if tid not in current_track_ids:
                roi_entry_times.pop(tid)

        # Handle Alerts
        if detection_alert and (time.time() - last_alert_time > ALERT_COOLDOWN):
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"alerts/alert_{timestamp}.jpg"
            cv2.imwrite(filename, frame)
            print(f"Alert saved: {filename}")
            
            send_telegram_alert(filename, alert_message)
            last_alert_time = time.time()
            
            # Visual feedback on screen
            cv2.putText(frame, "ALERT SAVED!", (50, 150),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

        # Global Alert Text
        if detection_alert:
            cv2.putText(frame, "SUSPICIOUS BEHAVIOR DETECTED!", (50, 100),
                        cv2.FONT_HERSHEY_SIMPLEX, 1.2, (0, 0, 255), 3)

        fps = 1.0 / (time.time() - start_time)
        cv2.putText(frame, f"FPS: {fps:.2f}", (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 0), 2)
        
        cv2.imshow("Theft Detection System", frame)

        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
