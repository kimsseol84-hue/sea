import React, { useEffect, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import { HandLandmarker, FilesetResolver } from "@mediapipe/tasks-vision";
import "./style.css";

const SEA_ANIMALS = ["🐠", "🐟", "🐙", "🦀", "🦑", "🐬", "🐡", "🪼", "🐳", "🦞"];
const MESSAGES = ["잡았다!", "좋아!", "성공!", "첨벙!"];

function makeAnimal(width, delay = 0) {
  const size = 56 + Math.random() * 34;
  return {
    emoji: SEA_ANIMALS[Math.floor(Math.random() * SEA_ANIMALS.length)],
    x: 60 + Math.random() * Math.max(100, width - 120),
    y: -90 - delay - Math.random() * 160,
    size,
    speed: 1.8 + Math.random() * 2.4,
    sway: Math.random() * Math.PI * 2,
    drawX: 0
  };
}

function App() {
  const canvasRef = useRef(null);
  const videoRef = useRef(null);
  const handRef = useRef(null);
  const rafRef = useRef(null);
  const timerRef = useRef(null);
  const lastVideoTimeRef = useRef(-1);

  const gameRef = useRef({
    running: false,
    mode: "touch",
    score: 0,
    timeLeft: 30,
    pointer: null,
    animals: [],
    bubbles: []
  });

  const [screen, setScreen] = useState("start");
  const [score, setScore] = useState(0);
  const [timeLeft, setTimeLeft] = useState(30);
  const [finalScore, setFinalScore] = useState(0);
  const [message, setMessage] = useState("");
  const [status, setStatus] = useState("손 인식 모드는 HTTPS 또는 localhost에서 가장 안정적으로 실행됩니다.");
  const [loading, setLoading] = useState(false);

  function resize() {
    const canvas = canvasRef.current;
    if (!canvas) return;
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
  }

  useEffect(() => {
    resize();
    window.addEventListener("resize", resize);
    return () => {
      window.removeEventListener("resize", resize);
      clearInterval(timerRef.current);
      cancelAnimationFrame(rafRef.current);
      const stream = videoRef.current?.srcObject;
      if (stream) stream.getTracks().forEach((track) => track.stop());
    };
  }, []);

  function resetGame(mode) {
    const canvas = canvasRef.current;
    const width = canvas?.width || window.innerWidth;

    gameRef.current = {
      running: true,
      mode,
      score: 0,
      timeLeft: 30,
      pointer: null,
      animals: [makeAnimal(width, 0), makeAnimal(width, 140), makeAnimal(width, 280)],
      bubbles: []
    };

    setScore(0);
    setTimeLeft(30);
    setFinalScore(0);
  }

  function startTimer() {
    clearInterval(timerRef.current);
    timerRef.current = setInterval(() => {
      const game = gameRef.current;
      game.timeLeft -= 1;
      setTimeLeft(game.timeLeft);
      if (game.timeLeft <= 0) endGame();
    }, 1000);
  }

  function endGame() {
    const game = gameRef.current;
    game.running = false;
    clearInterval(timerRef.current);
    cancelAnimationFrame(rafRef.current);
    setFinalScore(game.score);
    setScreen("end");
    setStatus("놀이가 끝났어요.");
  }

  function showCatchMessage(text) {
    setMessage(text);
    setTimeout(() => setMessage(""), 350);
  }

  async function setupCamera() {
    if (!navigator.mediaDevices?.getUserMedia) {
      throw new Error("이 브라우저는 카메라 기능을 지원하지 않습니다.");
    }

    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: "user", width: { ideal: 1280 }, height: { ideal: 720 } },
      audio: false
    });

    const video = videoRef.current;
    video.srcObject = stream;

    await new Promise((resolve) => {
      video.onloadedmetadata = () => {
        video.play();
        resolve();
      };
    });
  }

  async function setupHands() {
    if (handRef.current) return;

    const vision = await FilesetResolver.forVisionTasks(
      "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.14/wasm"
    );

    handRef.current = await HandLandmarker.createFromOptions(vision, {
      baseOptions: {
        modelAssetPath:
          "https://storage.googleapis.com/mediapipe-models/hand_landmarker/hand_landmarker/float16/1/hand_landmarker.task",
        delegate: "GPU"
      },
      runningMode: "VIDEO",
      numHands: 2
    });
  }

  function updateHandPointer() {
    const video = videoRef.current;
    const canvas = canvasRef.current;
    const hand = handRef.current;
    if (!video || !canvas || !hand || video.readyState < 2) return;

    if (video.currentTime !== lastVideoTimeRef.current) {
      lastVideoTimeRef.current = video.currentTime;
      const results = hand.detectForVideo(video, performance.now());

      if (results.landmarks?.length > 0) {
        const tip = results.landmarks[0][8];
        gameRef.current.pointer = {
          x: (1 - tip.x) * canvas.width,
          y: tip.y * canvas.height
        };
      } else {
        gameRef.current.pointer = null;
      }
    }
  }

  function addBubbles(x, y) {
    const game = gameRef.current;
    for (let i = 0; i < 10; i++) {
      game.bubbles.push({
        x,
        y,
        r: 4 + Math.random() * 10,
        vx: -1.6 + Math.random() * 3.2,
        vy: -1.2 - Math.random() * 3,
        life: 40
      });
    }
  }

  function checkCatch() {
    const game = gameRef.current;
    const pointer = game.pointer;
    if (!pointer) return;

    for (const animal of game.animals) {
      const dx = pointer.x - animal.drawX;
      const dy = pointer.y - animal.y;
      const dist = Math.sqrt(dx * dx + dy * dy);

      if (dist < animal.size * 0.68) {
        game.score += 1;
        setScore(game.score);
        addBubbles(animal.drawX, animal.y);
        showCatchMessage(MESSAGES[Math.floor(Math.random() * MESSAGES.length)]);
        Object.assign(animal, makeAnimal(canvasRef.current.width, 80));
      }
    }
  }

  function drawAnimals(ctx, canvas, game) {
    for (const animal of game.animals) {
      animal.y += animal.speed;
      animal.sway += 0.025;
      animal.drawX = animal.x + Math.sin(animal.sway) * 20;

      if (animal.y > canvas.height + 100) {
        Object.assign(animal, makeAnimal(canvas.width, 80));
      }

      ctx.save();
      ctx.shadowColor = "rgba(0,0,0,.35)";
      ctx.shadowBlur = 8;
      ctx.shadowOffsetY = 4;
      ctx.font = `${animal.size}px serif`;
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(animal.emoji, animal.drawX, animal.y);
      ctx.restore();
    }
  }

  function drawBubbles(ctx, game) {
    ctx.save();
    for (const bubble of game.bubbles) {
      bubble.x += bubble.vx;
      bubble.y += bubble.vy;
      bubble.life -= 1;

      ctx.globalAlpha = Math.max(0, bubble.life / 40);
      ctx.beginPath();
      ctx.arc(bubble.x, bubble.y, bubble.r, 0, Math.PI * 2);
      ctx.strokeStyle = "white";
      ctx.lineWidth = 2;
      ctx.stroke();
    }
    ctx.restore();

    game.bubbles = game.bubbles.filter((bubble) => bubble.life > 0);
  }

  function drawPointer(ctx, pointer) {
    if (!pointer) return;

    ctx.save();
    ctx.beginPath();
    ctx.arc(pointer.x, pointer.y, 27, 0, Math.PI * 2);
    ctx.fillStyle = "rgba(255,238,90,.78)";
    ctx.fill();

    ctx.lineWidth = 5;
    ctx.strokeStyle = "rgba(255,255,255,.92)";
    ctx.stroke();

    ctx.font = "27px serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("✋", pointer.x, pointer.y + 1);
    ctx.restore();
  }

  function draw() {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    const game = gameRef.current;

    ctx.clearRect(0, 0, canvas.width, canvas.height);
    drawAnimals(ctx, canvas, game);
    drawBubbles(ctx, game);
    drawPointer(ctx, game.pointer);
    checkCatch();
  }

  function loop() {
    const game = gameRef.current;
    if (!game.running) return;

    if (game.mode === "camera") updateHandPointer();
    draw();

    rafRef.current = requestAnimationFrame(loop);
  }

  function begin(mode) {
    resetGame(mode);
    setScreen("play");
    startTimer();
    loop();
  }

  async function startCameraMode() {
    setLoading(true);
    setStatus("카메라와 손 인식 모델을 준비하고 있어요.");

    try {
      await setupCamera();
      await setupHands();
      setStatus("손 인식 모드입니다. 검지 손끝으로 바다생물을 잡아보세요.");
      begin("camera");
    } catch (error) {
      console.error(error);
      setStatus(`카메라 모드 실패: ${error.name || "Error"} / ${error.message || "권한 또는 실행 주소 문제"}`);
      alert("카메라가 막혔어요. 배포된 HTTPS 주소 또는 localhost에서 실행해 주세요. 터치 모드로는 바로 사용할 수 있어요.");
    } finally {
      setLoading(false);
    }
  }

  function startTouchMode() {
    setStatus("터치 모드입니다. 화면을 누르거나 마우스를 움직여 바다생물을 잡아보세요.");
    begin("touch");
  }

  function updatePointerFromEvent(event) {
    const game = gameRef.current;
    if (!game.running || game.mode !== "touch") return;

    const point = event.touches ? event.touches[0] : event;
    game.pointer = { x: point.clientX, y: point.clientY };
  }

  function goHome() {
    gameRef.current.running = false;
    clearInterval(timerRef.current);
    cancelAnimationFrame(rafRef.current);
    setScreen("start");
    setStatus("처음 화면입니다.");
  }

  function restart() {
    const previousMode = gameRef.current.mode;
    if (previousMode === "camera") startCameraMode();
    else startTouchMode();
  }

  return (
    <div
      className="game"
      onMouseMove={updatePointerFromEvent}
      onTouchStart={updatePointerFromEvent}
      onTouchMove={updatePointerFromEvent}
    >
      <video
        ref={videoRef}
        className={screen === "play" && gameRef.current.mode === "camera" ? "video show" : "video"}
        autoPlay
        playsInline
        muted
      />

      <canvas ref={canvasRef} className="canvas" />

      <div className="ui">
        <div className="badge">🐚 점수 {score}</div>
        <div className="badge">⏰ {timeLeft}초</div>
      </div>

      <div className={message ? "center-message show" : "center-message"}>{message}</div>

      {screen === "start" && (
        <div className="panel">
          <div className="card">
            <h1>🌊 바다생물 잡기</h1>
            <p>카메라 화면 위로 내려오는 바다생물을<br />손으로 콕! 잡아보세요.</p>
            <p className="small">바닷속 배경 없이 실제 교실 화면 위에 바다생물만 나타나요.</p>
            <div className="buttons">
              <button onClick={startCameraMode} disabled={loading}>
                {loading ? "📷 준비 중..." : "📷 손 인식으로 시작"}
              </button>
              <button className="secondary" onClick={startTouchMode}>
                👆 터치 모드로 시작
              </button>
            </div>
          </div>
        </div>
      )}

      {screen === "end" && (
        <div className="panel">
          <div className="card">
            <h1>🎉 놀이 끝!</h1>
            <p>우리가 만난 바다생물은<br /><strong>{finalScore}마리</strong>예요.</p>
            <div className="buttons">
              <button onClick={restart}>다시 하기</button>
              <button className="secondary" onClick={goHome}>처음으로</button>
            </div>
          </div>
        </div>
      )}

      <div className="status">{status}</div>
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
