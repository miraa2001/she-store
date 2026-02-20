import { useEffect, useId, useState } from "react";
import styled from "styled-components";

function extractImageFiles(dataTransfer) {
  if (!dataTransfer?.items?.length) return [];

  return Array.from(dataTransfer.items)
    .filter((item) => item.kind === "file" && item.type?.startsWith("image/"))
    .map((item) => item.getAsFile())
    .filter(Boolean);
}

export default function FileUploadDropzone({
  disabled = false,
  onFilesSelected,
  currentCount = 0,
  maxImages = 0
}) {
  const inputId = useId();
  const [isDragOver, setIsDragOver] = useState(false);

  useEffect(() => {
    const handleWindowPaste = (event) => {
      if (disabled) return;
      const files = extractImageFiles(event.clipboardData);
      if (!files.length) return;
      event.preventDefault();
      onFilesSelected?.(files);
    };

    window.addEventListener("paste", handleWindowPaste);
    return () => window.removeEventListener("paste", handleWindowPaste);
  }, [disabled, onFilesSelected]);

  const handleDrop = (event) => {
    event.preventDefault();
    if (disabled) return;
    setIsDragOver(false);

    const files = Array.from(event.dataTransfer?.files || []).filter((file) =>
      file.type?.startsWith("image/")
    );
    if (files.length) onFilesSelected?.(files);
  };

  return (
    <StyledWrapper>
      <div className="file-upload-container">
        <div
          className={`file-upload ${isDragOver ? "dragover" : ""} ${disabled ? "disabled" : ""}`}
          onDragOver={(event) => {
            event.preventDefault();
            if (!disabled) setIsDragOver(true);
          }}
          onDragLeave={() => setIsDragOver(false)}
          onDrop={handleDrop}
        >
          <input
            multiple
            className="file-input"
            id={inputId}
            type="file"
            accept="image/*"
            onChange={(event) => {
              onFilesSelected?.(event.target.files);
              event.target.value = "";
            }}
            disabled={disabled}
          />
          <label className="file-label" htmlFor={inputId}>
            <i className="upload-icon" aria-hidden="true">
              📁
            </i>
            <p>Drag &amp; Drop الصور هنا أو اضغطي للرفع</p>
            <small>
              أو الصقي الصور مباشرة (Ctrl+V) — {currentCount}/{maxImages}
            </small>
          </label>
        </div>
      </div>
    </StyledWrapper>
  );
}

const StyledWrapper = styled.div`
  .file-upload-container {
    width: 100%;
  }

  .file-upload {
    position: relative;
    border: 2px dashed rgba(109, 63, 107, 0.38);
    border-radius: 12px;
    padding: 24px 16px;
    text-align: center;
    background-color: #ffffff;
    transition: background-color 0.2s ease, border-color 0.2s ease;
  }

  .file-upload:hover {
    background-color: rgba(109, 63, 107, 0.06);
    border-color: rgba(109, 63, 107, 0.62);
  }

  .file-upload.dragover {
    background-color: rgba(109, 63, 107, 0.14);
    border-color: #6D3F6B;
  }

  .file-upload.disabled {
    opacity: 0.55;
    pointer-events: none;
  }

  .file-input {
    display: none;
  }

  .file-label {
    display: flex;
    flex-direction: column;
    align-items: center;
    cursor: pointer;
    gap: 6px;
  }

  .upload-icon {
    font-size: 36px;
    color: #6D3F6B;
    margin-bottom: 2px;
    font-style: normal;
  }

  .file-upload p {
    margin: 0;
    font-size: 14px;
    color: #334155;
    font-weight: 600;
    line-height: 1.45;
  }

  .file-upload small {
    font-size: 12px;
    color: #64748b;
  }
`;
