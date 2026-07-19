import { useEffect, useMemo, useState } from "react";
import { DocumentViewer } from "../../src";
import type { DocumentSource } from "../../src";
import { blobSamples } from "./blobs";

export function App() {
  const [file, setFile] = useState<File | null>(null);
  const [selectedSampleId, setSelectedSampleId] = useState(blobSamples[0].id);
  const [sampleSource, setSampleSource] = useState<DocumentSource | null>(null);
  const [sampleError, setSampleError] = useState<Error | null>(null);
  const selectedSample = useMemo(
    () =>
      blobSamples.find((sample) => sample.id === selectedSampleId) ??
      blobSamples[0],
    [selectedSampleId],
  );

  useEffect(() => {
    let cancelled = false;

    setSampleSource(null);
    setSampleError(null);

    Promise.resolve(selectedSample.createSource()).then(
      (source) => {
        if (!cancelled) {
          setSampleSource(source);
        }
      },
      (error: unknown) => {
        if (!cancelled) {
          setSampleError(
            error instanceof Error
              ? error
              : new Error("Unable to create Blob sample."),
          );
        }
      },
    );

    return () => {
      cancelled = true;
    };
  }, [selectedSample]);

  const source = file ?? sampleSource;

  return (
    <main className="example-shell">
      <header className="example-header">
        <div>
          <h1>Local Document Viewer</h1>
          <p>
            Test hardcoded Blob documents or pick a file from this machine. The
            browser passes the data directly to the React component.
          </p>
        </div>
        <label className="example-file-button">
          <input
            onChange={(event) => {
              setFile(event.target.files?.[0] ?? null);
            }}
            type="file"
          />
          Select file
        </label>
      </header>

      <section className="example-samples" aria-label="Hardcoded Blob samples">
        <div className="example-sample-buttons">
          {blobSamples.map((sample) => (
            <button
              aria-pressed={!file && selectedSampleId === sample.id}
              className="example-sample-button"
              key={sample.id}
              onClick={() => {
                setFile(null);
                setSelectedSampleId(sample.id);
              }}
              type="button"
            >
              {sample.label}
            </button>
          ))}
        </div>
        <p>
          {file
            ? `Viewing uploaded file: ${file.name}`
            : selectedSample.description}
        </p>
      </section>

      {sampleError ? (
        <div className="example-error" role="alert">
          {sampleError.message}
        </div>
      ) : (
        <DocumentViewer
          height="calc(100vh - 204px)"
          onError={(error) => {
            console.error(error);
          }}
          pdfOptions={{ showThumbnails: true }}
          source={source}
          controls={{
            toolbar: true,
            fileName: true,
            pageNavigation: true,
            zoom: true,
            fit: true,
            rotate: true,
            search: true,
            print: true,
            download: true,
            fullscreen: true,
            thumbnails: true,
          }}
        />
      )}
    </main>
  );
}
