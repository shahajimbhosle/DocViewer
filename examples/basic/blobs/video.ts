import type { BlobSample } from "./types";

function supportedVideoMimeType(): string {
  const candidates = [
    "video/webm;codecs=vp9",
    "video/webm;codecs=vp8",
    "video/webm",
  ];

  return candidates.find((candidate) => MediaRecorder.isTypeSupported(candidate)) ?? "video/webm";
}

async function createVideoBlob(): Promise<Blob> {
  if (!("MediaRecorder" in globalThis)) {
    throw new Error("This browser does not support generated video Blob samples.");
  }

  const canvas = document.createElement("canvas");
  canvas.width = 640;
  canvas.height = 360;
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("Unable to prepare the video sample canvas.");
  }

  const drawingContext = context;
  const stream = canvas.captureStream(12);
  const recorder = new MediaRecorder(stream, {
    mimeType: supportedVideoMimeType(),
  });
  const chunks: BlobPart[] = [];

  const done = new Promise<Blob>((resolve, reject) => {
    recorder.ondataavailable = (event) => {
      if (event.data.size > 0) {
        chunks.push(event.data);
      }
    };
    recorder.onerror = () => {
      reject(new Error("Unable to record the generated video Blob sample."));
    };
    recorder.onstop = () => {
      stream.getTracks().forEach((track) => track.stop());
      resolve(new Blob(chunks, { type: recorder.mimeType || "video/webm" }));
    };
  });

  function drawFrame(progress: number) {
    const x = 80 + progress * 420;

    drawingContext.fillStyle = "#0f172a";
    drawingContext.fillRect(0, 0, canvas.width, canvas.height);
    drawingContext.fillStyle = "#14b8a6";
    drawingContext.fillRect(0, 0, canvas.width * progress, 12);
    drawingContext.fillStyle = "#ffffff";
    drawingContext.font = "bold 34px Arial";
    drawingContext.fillText("Local Video Blob", 52, 88);
    drawingContext.font = "20px Arial";
    drawingContext.fillText("Generated in the browser and passed to the viewer.", 52, 126);
    drawingContext.fillStyle = "#f97316";
    drawingContext.beginPath();
    drawingContext.arc(x, 230, 42, 0, Math.PI * 2);
    drawingContext.fill();
  }

  recorder.start();
  const startedAt = performance.now();

  function tick(now: number) {
    const progress = Math.min((now - startedAt) / 1000, 1);
    drawFrame(progress);

    if (progress < 1) {
      requestAnimationFrame(tick);
      return;
    }

    recorder.stop();
  }

  requestAnimationFrame(tick);

  return done;
}

export const videoBlobSample: BlobSample = {
  id: "video",
  label: "Video Blob",
  description: "Generated WebM video passed as a local video Blob.",
  createSource: async () => {
    const blob = await createVideoBlob();

    return {
      blob,
      fileName: "hardcoded-video.webm",
      mimeType: blob.type || "video/webm",
    };
  },
};
