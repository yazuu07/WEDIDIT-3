<?php
session_start();
include 'db.php';

$user_id = $_SESSION['user_id'];

if (!isset($user_id)) {
    echo json_encode(["success" => false, "error" => "Unauthorized access"]);
    exit();
}

$data = json_decode(file_get_contents("php://input"), true);

if (isset($data['image']) && isset($data['location'])) {
    $image_data = $data['image'];
    $location = $data['location'];

    $image_path = 'uploads/' . time() . '.jpg';
    $image_data = str_replace('data:image/jpeg;base64,', '', $image_data);
    $image_data = base64_decode($image_data);

    if (file_put_contents($image_path, $image_data)) {
        $query = "INSERT INTO uploads (user_id, image_path, location) VALUES (?, ?, ?)";
        $stmt = $conn->prepare($query);
        $stmt->bind_param("iss", $user_id, $image_path, $location);

        if ($stmt->execute()) {
            echo json_encode(["success" => true, "location" => $location]);
        } else {
            echo json_encode(["success" => false, "error" => "Database error"]);
        }
    } else {
        echo json_encode(["success" => false, "error" => "Failed to save image"]);
    }
} else {
    echo json_encode(["success" => false, "error" => "Invalid data"]);
}
?>

<?php
session_start();
if (!isset($_SESSION['user_id'])) {
    header("Location: login.php");
    exit();
}
?>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Photo Capture</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: Arial, sans-serif;
            background: #000;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            height: 100vh;
            width: 100vw;
            overflow: hidden;
        }
        .camera-container {
            position: relative;
            width: 100vw;
            height: 100vh;
            display: flex;
            justify-content: center;
            align-items: center;
            background: #222;
        }
        video, img {
            width: 100%;
            height: 100%;
            object-fit: cover;
            transform: scaleX(-1);
        }
        video { position: relative; }
        canvas { display: none; }
        .buttons {
            position: absolute;
            bottom: 20px;
            display: flex;
            gap: 10px;
        }
       /* Bottom Frame */
.bottom-frame {
    position: absolute;
    height: 100px;
    bottom: 0;
    left: 0;
    width: 100%;
    background: rgba(0, 0, 0, 0.7);
    padding: 15px 0;
    display: flex;
    justify-content: center;
    align-items: center;
    border-top-left-radius: 15px;
    border-top-right-radius: 15px;
}

/* Button Layout */
.buttons {
    display: flex;
    gap: 15px;
}

/* Circular Buttons */
button {
    width: 60px;
    height: 60px;
    border-radius: 50%;
    border: none;
    cursor: pointer;
    font-size: 12px;
    font-weight: bold;
    color: white;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.3s ease;
}

/* Button Colors */
.capture-btn { background: #ffffff; color: black; }
.retake-btn { background:rgb(255, 255, 255);color: black; }
.confirm-btn { background: rgb(255, 255, 255);color: black; }
.undertime-btn { background: rgb(255, 255, 255);color: black;}
.overtime-btn { background:rgb(255, 255, 255);color: black; }
.out-btn { background: rgb(255, 255, 255); color: black;}

button:hover {
    opacity: 0.8;
}


    </style>
</head>
<body>
    <div class="camera-container">
        <video id="video" autoplay></video>
        <canvas id="canvas"></canvas>
        <img id="photo" style="display:none;">
    </div>

    <div class="bottom-frame">
    <div class="buttons">
        <button id="capture" class="capture-btn">Capture</button>
        <button id="retake" class="retake-btn" style="display: none;">Retake</button>
        <button id="confirm" class="confirm-btn" style="display: none;">Time In</button>
        <button id="undertime" class="undertime-btn" style="display: none;">Undertime</button>
        <button id="overtime" class="overtime-btn" style="display: none;">Overtime</button>
        <button id="out" class="out-btn" style="display: none;">Time Out</button>
    </div>
</div>


    <script>
        const video = document.getElementById("video");
        const canvas = document.getElementById("canvas");
        const ctx = canvas.getContext("2d");
        const photo = document.getElementById("photo");
        const captureBtn = document.getElementById("capture");
        const retakeBtn = document.getElementById("retake");
        const confirmBtn = document.getElementById("confirm");
        const undertimeBtn = document.getElementById("undertime");
        const overtimeBtn = document.getElementById("overtime");
        const outBtn = document.getElementById("out");

        let lastPhotoData = null;
        let locationData = "Fetching location...";

        navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" } })
            .then(stream => video.srcObject = stream)
            .catch(err => console.error("Camera access denied:", err));

        navigator.geolocation.getCurrentPosition(
            async (position) => {
                const { latitude, longitude } = position.coords;
                try {
                    const response = await fetch(
                        `https://nominatim.openstreetmap.org/reverse?format=json&lat=${latitude}&lon=${longitude}`
                    );
                    const data = await response.json();
                    locationData = data.display_name || "Unknown Location";
                } catch (error) {
                    locationData = "Location fetch failed";
                }
            },
            () => locationData = "Location access denied"
        );

        captureBtn.addEventListener("click", () => {
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

            ctx.fillStyle = "white";
            ctx.font = "15px Arial";
            const dateTime = new Date().toLocaleString();
            ctx.fillText(dateTime, 20, canvas.height - 90);
            ctx.fillText(locationData, 20, canvas.height - 50);

            lastPhotoData = canvas.toDataURL("image/jpeg");
            photo.src = lastPhotoData;
            photo.style.display = "block";
            video.style.display = "none";
            captureBtn.style.display = "none";
            retakeBtn.style.display = "block";
            confirmBtn.style.display = "block";
            undertimeBtn.style.display = "block";
            overtimeBtn.style.display = "block";
            outBtn.style.display = "block";
        });

        retakeBtn.addEventListener("click", () => {
            photo.style.display = "none";
            video.style.display = "block";
            captureBtn.style.display = "block";
            retakeBtn.style.display = "none";
            confirmBtn.style.display = "none";
            undertimeBtn.style.display = "none";
            overtimeBtn.style.display = "none";
            outBtn.style.display = "none";
        });

        function uploadImage(locationType) {
            if (!lastPhotoData) {
                alert("Please capture an image first!");
                return;
            }

            fetch("upload_image.php", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ image: lastPhotoData, location: locationType }),
            })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    alert(`Image uploaded successfully as ${locationType}!`);
                    retakeBtn.click();
                } else {
                    alert("Upload failed: " + data.error);
                }
            })
            .catch(error => alert("Error uploading image."));
        }

        confirmBtn.addEventListener("click", () => uploadImage("In"));
        undertimeBtn.addEventListener("click", () => uploadImage("Undertime"));
        overtimeBtn.addEventListener("click", () => uploadImage("Overtime"));
        outBtn.addEventListener("click", () => uploadImage("Out"));
    </script>
</body>
</html>
