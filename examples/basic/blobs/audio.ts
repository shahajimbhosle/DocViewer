import type { BlobSample } from "./types";

function createWavBlob(): Blob {
  const sampleRate = 44100;
  const durationSeconds = 1;
  const sampleCount = sampleRate * durationSeconds;
  const bytesPerSample = 2;
  const dataSize = sampleCount * bytesPerSample;
  const buffer = new ArrayBuffer(44 + dataSize);
  const view = new DataView(buffer);

  function writeString(offset: number, value: string) {
    for (let index = 0; index < value.length; index += 1) {
      view.setUint8(offset + index, value.charCodeAt(index));
    }
  }

  writeString(0, "RIFF");
  view.setUint32(4, 36 + dataSize, true);
  writeString(8, "WAVE");
  writeString(12, "fmt ");
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, sampleRate, true);
  view.setUint32(28, sampleRate * bytesPerSample, true);
  view.setUint16(32, bytesPerSample, true);
  view.setUint16(34, 16, true);
  writeString(36, "data");
  view.setUint32(40, dataSize, true);

  for (let index = 0; index < sampleCount; index += 1) {
    const fade = Math.min(index / 3000, (sampleCount - index) / 3000, 1);
    const value = Math.sin((2 * Math.PI * 440 * index) / sampleRate) * 0.28 * fade;
    view.setInt16(44 + index * bytesPerSample, value * 32767, true);
  }

  return new Blob([buffer], { type: "audio/wav" });
}

export const audioBlobSample: BlobSample = {
  id: "audio",
  label: "Audio Blob",
  description: "Generated WAV audio passed as a local audio/wav Blob.",
  createSource: () => ({
    blob: createWavBlob(),
    fileName: "hardcoded-tone.wav",
    mimeType: "audio/wav",
  }),
};
