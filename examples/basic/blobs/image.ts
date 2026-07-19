import type { BlobSample } from "./types";

async function createPngBlob(): Promise<Blob> {
  const canvas = document.createElement("canvas");
  canvas.width = 640;
  canvas.height = 360;
  const context = canvas.getContext("2d");

  if (!context) {
    throw new Error("Unable to prepare the image sample canvas.");
  }

  const gradient = context.createLinearGradient(0, 0, canvas.width, canvas.height);
  gradient.addColorStop(0, "#0f766e");
  gradient.addColorStop(0.55, "#2563eb");
  gradient.addColorStop(1, "#f97316");
  context.fillStyle = gradient;
  context.fillRect(0, 0, canvas.width, canvas.height);

  context.fillStyle = "rgba(255, 255, 255, 0.16)";
  context.beginPath();
  context.arc(520, 68, 90, 0, Math.PI * 2);
  context.fill();
  context.beginPath();
  context.arc(112, 304, 130, 0, Math.PI * 2);
  context.fill();

  context.fillStyle = "#ffffff";
  context.font = "bold 44px Arial";
  context.fillText("Local Image Blob", 64, 142);
  context.font = "22px Arial";
  context.fillText("Generated in this browser and rendered locally.", 64, 186);

  context.fillStyle = "rgba(15, 23, 42, 0.72)";
  context.fillRect(64, 238, 306, 58);
  context.fillStyle = "#ffffff";
  context.font = "bold 24px Arial";
  context.fillText("image/png", 92, 276);

  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (blob) {
          resolve(blob);
          return;
        }

        reject(new Error("Unable to create the PNG Blob sample."));
      },
      "image/png",
      0.92,
    );
  });
}

export const imageBlobSample: BlobSample = {
  id: "image",
  label: "Image Blob",
  description: "PNG image bytes passed as a local image/png Blob.",
  createSource: async () => ({
    blob: await createPngBlob(),
    fileName: "hardcoded-image.png",
  }),
};
