import { useEffect, useRef, useState } from "react";
import { fabric } from "fabric";
import "./ImageAnnotatorModal.css";

const COLORS = [
  "#FF0000",
  "#00FF00",
  "#0000FF",
  "#FFFF00",
  "#FF00FF",
  "#00FFFF",
  "#000000",
  "#FFFFFF",
  "#FFA500",
  "#800080"
];

const BRUSH_SIZES = {
  thin: 2,
  medium: 5,
  thick: 10
};

const TEXT_SIZES = {
  thin: 16,
  medium: 24,
  thick: 32
};

const TOOL_OPTIONS = [
  { id: "brush", label: "\u0641\u0631\u0634\u0627\u0629", hint: "Brush" },
  { id: "rectangle", label: "\u0645\u0633\u062A\u0637\u064A\u0644", hint: "Rectangle" },
  { id: "crop", label: "\u0642\u0635", hint: "Crop (handles)" },
  { id: "text", label: "\u0646\u0635", hint: "Text" }
];

const SIZE_OPTIONS = [
  { id: "thin", label: "\u0631\u0641\u064A\u0639", hint: "Thin" },
  { id: "medium", label: "\u0645\u062A\u0648\u0633\u0637", hint: "Medium" },
  { id: "thick", label: "\u0633\u0645\u064A\u0643", hint: "Thick" }
];

function isTextObject(object) {
  if (!object) return false;
  return object.type === "text" || object.type === "i-text" || object.type === "textbox";
}

function toEditedFileName(name) {
  const base = String(name || "image").replace(/\.[^/.]+$/, "") || "image";
  return `${base}-edited.png`;
}

export default function ImageAnnotatorModal({
  open,
  file,
  onCancel,
  onClose,
  onSave,
  disabled = false
}) {
  const canvasRef = useRef(null);
  const fabricCanvasRef = useRef(null);
  const imageUrlRef = useRef(null);
  const sourceImageRef = useRef(null);

  const [isLoading, setIsLoading] = useState(true);
  const [loadError, setLoadError] = useState("");
  const [saving, setSaving] = useState(false);
  const [originalDimensions, setOriginalDimensions] = useState({ width: 0, height: 0 });
  const [displayDimensions, setDisplayDimensions] = useState({ width: 0, height: 0 });
  const [canvasReadyTick, setCanvasReadyTick] = useState(0);
  const [previewUrl, setPreviewUrl] = useState("");
  const [activeTool, setActiveTool] = useState("brush");
  const [brushSize, setBrushSize] = useState("medium");
  const [selectedColor, setSelectedColor] = useState("#FF0000");
  const [showColorPicker, setShowColorPicker] = useState(false);
  const [applyingCrop, setApplyingCrop] = useState(false);
  const [textDraft, setTextDraft] = useState("");

  const closeHandler = onCancel || onClose;

  const computeDisplayDimensions = (width, height) => {
    const maxWidth = Math.min(window.innerWidth * 0.85, 1200);
    const maxHeight = Math.min(window.innerHeight * 0.6, 800);
    const scale = Math.min(maxWidth / width, maxHeight / height, 1);
    return {
      scale,
      width: Math.max(1, Math.floor(width * scale)),
      height: Math.max(1, Math.floor(height * scale))
    };
  };

  const extractCropRegion = (canvas) => {
    if (!canvas) return null;
    const cropObject = canvas.getObjects().find((obj) => obj?.isCrop);
    if (!cropObject) return null;

    const scaleFactor = 1 / canvas.originalScale;
    const rawX = Number(cropObject.left || 0) * scaleFactor;
    const rawY = Number(cropObject.top || 0) * scaleFactor;
    const rawW =
      Number(
        typeof cropObject.getScaledWidth === "function"
          ? cropObject.getScaledWidth()
          : (cropObject.width || 0) * (cropObject.scaleX || 1)
      ) * scaleFactor;
    const rawH =
      Number(
        typeof cropObject.getScaledHeight === "function"
          ? cropObject.getScaledHeight()
          : (cropObject.height || 0) * (cropObject.scaleY || 1)
      ) * scaleFactor;

    const x = Math.max(0, Math.min(originalDimensions.width - 1, Math.round(rawX)));
    const y = Math.max(0, Math.min(originalDimensions.height - 1, Math.round(rawY)));
    const w = Math.max(1, Math.round(rawW));
    const h = Math.max(1, Math.round(rawH));
    const clampedW = Math.max(1, Math.min(w, originalDimensions.width - x));
    const clampedH = Math.max(1, Math.min(h, originalDimensions.height - y));
    return { x, y, w: clampedW, h: clampedH };
  };

  const renderMergedCanvas = async (canvas) => {
    if (!canvas) throw new Error("Missing canvas");
    if (!sourceImageRef.current) throw new Error("Missing source image");

    const mergedCanvas = document.createElement("canvas");
    mergedCanvas.width = originalDimensions.width;
    mergedCanvas.height = originalDimensions.height;
    const mergedCtx = mergedCanvas.getContext("2d");
    if (!mergedCtx) throw new Error("2D canvas context unavailable");

    mergedCtx.drawImage(sourceImageRef.current, 0, 0, originalDimensions.width, originalDimensions.height);

    const tempCanvas = new fabric.Canvas(null, {
      width: originalDimensions.width,
      height: originalDimensions.height
    });

    const scaleFactor = 1 / canvas.originalScale;
    const objects = canvas.getObjects().filter((obj) => !obj?.isCrop);
    for (const obj of objects) {
      // eslint-disable-next-line no-await-in-loop
      const cloned = await new Promise((resolve) => {
        obj.clone((clonedObj) => {
          clonedObj.set({
            scaleX: (obj.scaleX || 1) * scaleFactor,
            scaleY: (obj.scaleY || 1) * scaleFactor,
            left: (obj.left || 0) * scaleFactor,
            top: (obj.top || 0) * scaleFactor,
            strokeWidth: (obj.strokeWidth || 1) * scaleFactor
          });

          if (clonedObj.type === "path" && clonedObj.path) {
            clonedObj.path = clonedObj.path.map((pathPoint) =>
              pathPoint.map((val, idx) => (idx > 0 && typeof val === "number" ? val * scaleFactor : val))
            );
          }
          resolve(clonedObj);
        });
      });
      tempCanvas.add(cloned);
    }

    tempCanvas.requestRenderAll();
    const drawingsDataUrl = tempCanvas.toDataURL({ format: "png", quality: 1, multiplier: 1 });
    tempCanvas.dispose();

    const drawingsImage = new Image();
    await new Promise((resolve, reject) => {
      drawingsImage.onload = resolve;
      drawingsImage.onerror = reject;
      drawingsImage.src = drawingsDataUrl;
    });
    mergedCtx.drawImage(drawingsImage, 0, 0);

    return mergedCanvas;
  };

  const finalizeTextEditing = () => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    canvas.getObjects().forEach((obj) => {
      if (isTextObject(obj) && typeof obj.exitEditing === "function" && obj.isEditing) {
        obj.exitEditing();
      }
    });
    canvas.discardActiveObject();
    canvas.requestRenderAll();
  };

  const applyCropSelection = async () => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return false;

    const cropRegion = extractCropRegion(canvas);
    if (!cropRegion) return false;

    const isFullImageCrop =
      cropRegion.x === 0 &&
      cropRegion.y === 0 &&
      cropRegion.w >= originalDimensions.width &&
      cropRegion.h >= originalDimensions.height;

    if (isFullImageCrop) {
      canvas.getObjects().filter((obj) => obj?.isCrop).forEach((obj) => canvas.remove(obj));
      canvas.requestRenderAll();
      return false;
    }

    try {
      setApplyingCrop(true);
      setIsLoading(true);

      const mergedCanvas = await renderMergedCanvas(canvas);
      const croppedCanvas = document.createElement("canvas");
      croppedCanvas.width = cropRegion.w;
      croppedCanvas.height = cropRegion.h;
      const croppedCtx = croppedCanvas.getContext("2d");
      if (!croppedCtx) throw new Error("Crop canvas context unavailable");

      croppedCtx.drawImage(
        mergedCanvas,
        cropRegion.x,
        cropRegion.y,
        cropRegion.w,
        cropRegion.h,
        0,
        0,
        cropRegion.w,
        cropRegion.h
      );

      const dataUrl = croppedCanvas.toDataURL("image/png", 1);
      const nextImage = new Image();
      await new Promise((resolve, reject) => {
        nextImage.onload = resolve;
        nextImage.onerror = reject;
        nextImage.src = dataUrl;
      });

      if (fabricCanvasRef.current) {
        fabricCanvasRef.current.dispose();
        fabricCanvasRef.current = null;
      }

      const nextOriginalWidth = nextImage.naturalWidth || nextImage.width;
      const nextOriginalHeight = nextImage.naturalHeight || nextImage.height;
      const nextDisplay = computeDisplayDimensions(nextOriginalWidth, nextOriginalHeight);

      setOriginalDimensions({ width: nextOriginalWidth, height: nextOriginalHeight });
      setDisplayDimensions({ width: nextDisplay.width, height: nextDisplay.height });
      setPreviewUrl(dataUrl);
      sourceImageRef.current = nextImage;

      const nextCanvas = new fabric.Canvas(canvasRef.current, {
        width: nextDisplay.width,
        height: nextDisplay.height,
        backgroundColor: "transparent",
        preserveObjectStacking: true,
        selection: false
      });
      nextCanvas.originalScale = nextDisplay.scale;
      nextCanvas.originalWidth = nextOriginalWidth;
      nextCanvas.originalHeight = nextOriginalHeight;
      fabricCanvasRef.current = nextCanvas;
      setCanvasReadyTick((prev) => prev + 1);
      setIsLoading(false);
      return true;
    } catch (error) {
      setLoadError(error?.message || "Failed to apply crop");
      setIsLoading(false);
      return false;
    } finally {
      setApplyingCrop(false);
    }
  };

  const handleToolChange = async (nextTool) => {
    if (nextTool === activeTool) return;

    if (activeTool === "text") {
      finalizeTextEditing();
    }

    if (activeTool === "crop" && nextTool !== "crop") {
      await applyCropSelection();
    }

    setActiveTool(nextTool);
  };

  const handleSizeChange = (nextSize) => {
    setBrushSize(nextSize);
    if (activeTool !== "text") return;
    const canvas = fabricCanvasRef.current;
    const activeObject = canvas?.getActiveObject?.();
    if (!activeObject || !isTextObject(activeObject)) return;
    activeObject.set("fontSize", TEXT_SIZES[nextSize] || TEXT_SIZES.medium);
    canvas.requestRenderAll();
  };

  useEffect(() => {
    if (!["brush", "rectangle", "crop", "text"].includes(activeTool)) {
      setActiveTool("brush");
    }
  }, [activeTool]);

  useEffect(() => {
    if (!open) return;
    setActiveTool("brush");
    setBrushSize("medium");
    setSelectedColor("#FF0000");
    setTextDraft("");
  }, [open]);

  useEffect(() => {
    if (!open || !file) return undefined;

    let mounted = true;
    setIsLoading(true);
    setLoadError("");

    const imageUrl = URL.createObjectURL(file);
    imageUrlRef.current = imageUrl;
    setPreviewUrl(imageUrl);

    const img = new Image();
    img.onload = () => {
      if (!mounted) return;

      const originalWidth = img.naturalWidth || img.width;
      const originalHeight = img.naturalHeight || img.height;

      if (!originalWidth || !originalHeight) {
        setLoadError("Invalid image dimensions");
        setIsLoading(false);
        return;
      }

      const nextDisplay = computeDisplayDimensions(originalWidth, originalHeight);
      const scale = nextDisplay.scale;
      const displayWidth = nextDisplay.width;
      const displayHeight = nextDisplay.height;

      setOriginalDimensions({ width: originalWidth, height: originalHeight });
      setDisplayDimensions({ width: displayWidth, height: displayHeight });
      sourceImageRef.current = img;

      if (fabricCanvasRef.current) {
        fabricCanvasRef.current.dispose();
        fabricCanvasRef.current = null;
      }

      if (!canvasRef.current) {
        setLoadError("Canvas element not found");
        setIsLoading(false);
        return;
      }

      const canvas = new fabric.Canvas(canvasRef.current, {
        width: displayWidth,
        height: displayHeight,
        backgroundColor: "transparent",
        preserveObjectStacking: true,
        selection: false
      });

      canvas.originalScale = scale;
      canvas.originalWidth = originalWidth;
      canvas.originalHeight = originalHeight;
      fabricCanvasRef.current = canvas;
      setCanvasReadyTick((prev) => prev + 1);

      canvas.requestRenderAll();
      setIsLoading(false);
    };

    img.onerror = () => {
      if (!mounted) return;
      setLoadError("Failed to load image");
      setIsLoading(false);
    };

    img.src = imageUrl;

    return () => {
      mounted = false;
      sourceImageRef.current = null;
      setPreviewUrl("");
      if (fabricCanvasRef.current) {
        fabricCanvasRef.current.dispose();
        fabricCanvasRef.current = null;
      }
      if (imageUrlRef.current) {
        URL.revokeObjectURL(imageUrlRef.current);
        imageUrlRef.current = null;
      }
    };
  }, [open, file]);

  useEffect(() => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;

    canvas.off("path:created");
    canvas.off("object:scaling");
    canvas.off("object:modified");
    canvas.off("mouse:down");
    canvas.off("mouse:move");
    canvas.off("mouse:up");

    if (activeTool === "brush") {
      canvas.isDrawingMode = true;
      canvas.selection = false;

      const brush = new fabric.PencilBrush(canvas);
      brush.width = BRUSH_SIZES[brushSize] || BRUSH_SIZES.medium;
      brush.color = selectedColor;
      canvas.freeDrawingBrush = brush;

      canvas.on("path:created", (event) => {
        const path = event?.path;
        if (!path) return;
        path.set({
          selectable: false,
          evented: false,
          hasControls: false,
          hasBorders: false,
          lockMovementX: true,
          lockMovementY: true,
          lockRotation: true
        });
        canvas.requestRenderAll();
      });
      return;
    }

    if (activeTool === "rectangle") {
      canvas.isDrawingMode = false;
      canvas.selection = false;

      let rect = null;
      let isDown = false;
      let startX = 0;
      let startY = 0;

      canvas.on("mouse:down", (event) => {
        isDown = true;
        const pointer = canvas.getPointer(event.e);
        startX = pointer.x;
        startY = pointer.y;

        rect = new fabric.Rect({
          left: startX,
          top: startY,
          width: 0,
          height: 0,
          fill: "transparent",
          stroke: selectedColor,
          strokeWidth: BRUSH_SIZES[brushSize] || BRUSH_SIZES.medium,
          selectable: false,
          evented: false,
          hasControls: false,
          hasBorders: false,
          lockMovementX: true,
          lockMovementY: true,
          lockRotation: true
        });
        canvas.add(rect);
      });

      canvas.on("mouse:move", (event) => {
        if (!isDown || !rect) return;
        const pointer = canvas.getPointer(event.e);
        rect.set({
          left: Math.min(startX, pointer.x),
          top: Math.min(startY, pointer.y),
          width: Math.abs(pointer.x - startX),
          height: Math.abs(pointer.y - startY)
        });
        canvas.requestRenderAll();
      });

      canvas.on("mouse:up", () => {
        isDown = false;
        if (rect && (!rect.width || !rect.height)) {
          canvas.remove(rect);
        }
        rect = null;
      });
      return;
    }

    if (activeTool === "crop") {
      canvas.isDrawingMode = false;
      canvas.selection = true;

      const allObjects = canvas.getObjects();
      allObjects.forEach((object) => {
        if (!object?.isCrop) {
          object.set({
            selectable: false,
            evented: false
          });
        }
      });

      const minCropSize = 20;
      let cropRect = allObjects.find((object) => object?.isCrop);
      if (!cropRect) {
        const margin = 12;
        cropRect = new fabric.Rect({
          left: margin,
          top: margin,
          width: Math.max(minCropSize, canvas.getWidth() - margin * 2),
          height: Math.max(minCropSize, canvas.getHeight() - margin * 2),
          fill: "rgba(255,255,255,0.12)",
          stroke: selectedColor,
          strokeWidth: 2,
          strokeDashArray: [8, 6],
          selectable: true,
          evented: true,
          hasControls: true,
          hasBorders: true,
          lockMovementX: false,
          lockMovementY: false,
          lockRotation: true,
          cornerSize: 20,
          touchCornerSize: 28,
          isCrop: true
        });
        cropRect.setControlsVisibility({
          tl: false,
          tr: false,
          bl: false,
          br: false,
          ml: true,
          mr: true,
          mt: true,
          mb: true,
          mtr: false
        });
        canvas.add(cropRect);
      }

      cropRect.set({
        selectable: true,
        evented: true,
        hasControls: true,
        hasBorders: true,
        lockMovementX: false,
        lockMovementY: false,
        lockRotation: true,
        stroke: selectedColor,
        borderColor: selectedColor,
        cornerColor: selectedColor,
        cornerStrokeColor: "#ffffff",
        cornerStyle: "circle",
        transparentCorners: false,
        minScaleLimit: 0.05,
        cornerSize: 20,
        touchCornerSize: 28
      });
      cropRect.setControlsVisibility({
        tl: false,
        tr: false,
        bl: false,
        br: false,
        ml: true,
        mr: true,
        mt: true,
        mb: true,
        mtr: false
      });
      canvas.setActiveObject(cropRect);
      canvas.requestRenderAll();
      return;
    }

    if (activeTool === "text") {
      canvas.isDrawingMode = false;
      canvas.selection = true;

      canvas.getObjects().forEach((object) => {
        if (isTextObject(object)) {
          object.set({
            selectable: true,
            evented: true,
            hasControls: true,
            hasBorders: true,
            lockScalingX: false,
            lockScalingY: false,
            lockRotation: true
          });
          object.setControlsVisibility({
            tl: true,
            tr: true,
            bl: true,
            br: true,
            ml: false,
            mr: false,
            mt: false,
            mb: false,
            mtr: false
          });
        } else if (!object?.isCrop) {
          object.set({
            selectable: false,
            evented: false
          });
        }
      });

      canvas.on("mouse:down", (event) => {
        if (isTextObject(event?.target)) {
          return;
        }

        const pointer = canvas.getPointer(event.e);
        const textObject = new fabric.Textbox(String(textDraft || "نص"), {
          left: pointer.x,
          top: pointer.y,
          width: 220,
          fill: selectedColor,
          fontSize: TEXT_SIZES[brushSize] || TEXT_SIZES.medium,
          fontFamily: "Arial",
          editable: true,
          selectable: true,
          evented: true,
          hasControls: true,
          hasBorders: true,
          lockScalingX: false,
          lockScalingY: false,
          lockRotation: true
        });
        textObject.setControlsVisibility({
          tl: true,
          tr: true,
          bl: true,
          br: true,
          ml: false,
          mr: false,
          mt: false,
          mb: false,
          mtr: false
        });
        canvas.add(textObject);
        canvas.setActiveObject(textObject);
        canvas.requestRenderAll();
      });
      return;
    }

  }, [activeTool, brushSize, selectedColor, canvasReadyTick]);

  const handleUndo = () => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    const objects = canvas.getObjects();
    if (!objects.length) return;
    canvas.remove(objects[objects.length - 1]);
    canvas.requestRenderAll();
  };

  const handleClear = () => {
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;
    const objects = canvas.getObjects();
    objects.forEach((object) => canvas.remove(object));
    canvas.requestRenderAll();
  };

  const handleSave = async () => {
    if (saving || disabled) return;
    const canvas = fabricCanvasRef.current;
    if (!canvas) return;

    try {
      setSaving(true);
      setLoadError("");
      setIsLoading(true);

      const exportCanvas = document.createElement("canvas");
      exportCanvas.width = originalDimensions.width;
      exportCanvas.height = originalDimensions.height;
      const ctx = exportCanvas.getContext("2d");
      if (!ctx) throw new Error("2D canvas context unavailable");

      if (!sourceImageRef.current) {
        throw new Error("Missing source image");
      }
      ctx.drawImage(sourceImageRef.current, 0, 0, originalDimensions.width, originalDimensions.height);

      const tempCanvas = new fabric.Canvas(null, {
        width: originalDimensions.width,
        height: originalDimensions.height
      });

      const scaleFactor = 1 / canvas.originalScale;
      const cropObject = canvas.getObjects().find((obj) => obj?.isCrop);
      let cropRegion = null;
      if (cropObject) {
        const rawX = Number(cropObject.left || 0) * scaleFactor;
        const rawY = Number(cropObject.top || 0) * scaleFactor;
        const rawW =
          Number(
            typeof cropObject.getScaledWidth === "function"
              ? cropObject.getScaledWidth()
              : (cropObject.width || 0) * (cropObject.scaleX || 1)
          ) * scaleFactor;
        const rawH =
          Number(
            typeof cropObject.getScaledHeight === "function"
              ? cropObject.getScaledHeight()
              : (cropObject.height || 0) * (cropObject.scaleY || 1)
          ) * scaleFactor;

        const x = Math.max(0, Math.min(originalDimensions.width - 1, Math.round(rawX)));
        const y = Math.max(0, Math.min(originalDimensions.height - 1, Math.round(rawY)));
        const w = Math.max(1, Math.round(rawW));
        const h = Math.max(1, Math.round(rawH));
        const clampedW = Math.max(1, Math.min(w, originalDimensions.width - x));
        const clampedH = Math.max(1, Math.min(h, originalDimensions.height - y));
        cropRegion = { x, y, w: clampedW, h: clampedH };
      }

      const objects = canvas.getObjects().filter((obj) => !obj?.isCrop);
      for (const obj of objects) {
        // eslint-disable-next-line no-await-in-loop
        const cloned = await new Promise((resolve) => {
          obj.clone((clonedObj) => {
            clonedObj.set({
              scaleX: (obj.scaleX || 1) * scaleFactor,
              scaleY: (obj.scaleY || 1) * scaleFactor,
              left: (obj.left || 0) * scaleFactor,
              top: (obj.top || 0) * scaleFactor,
              strokeWidth: (obj.strokeWidth || 1) * scaleFactor
            });

            if (clonedObj.type === "path" && clonedObj.path) {
              clonedObj.path = clonedObj.path.map((pathPoint) =>
                pathPoint.map((val, idx) => (idx > 0 && typeof val === "number" ? val * scaleFactor : val))
              );
            }
            resolve(clonedObj);
          });
        });
        tempCanvas.add(cloned);
      }

      tempCanvas.requestRenderAll();
      const drawingsDataUrl = tempCanvas.toDataURL({ format: "png", quality: 1, multiplier: 1 });
      tempCanvas.dispose();

      const drawingsImage = new Image();
      await new Promise((resolve, reject) => {
        drawingsImage.onload = resolve;
        drawingsImage.onerror = reject;
        drawingsImage.src = drawingsDataUrl;
      });
      ctx.drawImage(drawingsImage, 0, 0);

      let outputCanvas = exportCanvas;
      if (cropRegion) {
        const croppedCanvas = document.createElement("canvas");
        croppedCanvas.width = cropRegion.w;
        croppedCanvas.height = cropRegion.h;
        const croppedCtx = croppedCanvas.getContext("2d");
        if (!croppedCtx) throw new Error("Crop canvas context unavailable");
        croppedCtx.drawImage(
          exportCanvas,
          cropRegion.x,
          cropRegion.y,
          cropRegion.w,
          cropRegion.h,
          0,
          0,
          cropRegion.w,
          cropRegion.h
        );
        outputCanvas = croppedCanvas;
      }

      const blob = await new Promise((resolve, reject) => {
        outputCanvas.toBlob(
          (outputBlob) => {
            if (!outputBlob) {
              reject(new Error("Failed to export edited image"));
              return;
            }
            resolve(outputBlob);
          },
          "image/png",
          1
        );
      });

      const editedFile = new File([blob], toEditedFileName(file.name), {
        type: "image/png",
        lastModified: Date.now()
      });

      setIsLoading(false);
      onSave?.(editedFile);
    } catch (error) {
      setIsLoading(false);
      setLoadError(error?.message || "Failed to save edited image");
    } finally {
      setSaving(false);
    }
  };

  if (!open) return null;

  return (
    <div
      className="image-annotator-overlay"
      onClick={(event) => event.stopPropagation()}
      onMouseDown={(event) => event.stopPropagation()}
      onPointerDown={(event) => event.stopPropagation()}
    >
      <div
        className="image-annotator-modal"
        onClick={(event) => event.stopPropagation()}
        onMouseDown={(event) => event.stopPropagation()}
        onPointerDown={(event) => event.stopPropagation()}
      >
        <div className="annotator-header">
          <h3>Edit image</h3>
          <button type="button" className="close-btn" onClick={closeHandler} disabled={disabled || saving || applyingCrop}>
            X
          </button>
        </div>

        <div className="annotator-toolbar">
          <div className="tool-section">
            <label>{`\u0627\u0644\u0623\u062F\u0627\u0629:`}</label>
            <div className="tool-buttons">
              {TOOL_OPTIONS.map((tool) => (
                <button
                  key={tool.id}
                  type="button"
                  className={activeTool === tool.id ? "active" : ""}
                  onClick={() => {
                    handleToolChange(tool.id);
                  }}
                  aria-label={tool.hint}
                  title={tool.hint}
                  disabled={isLoading || saving || applyingCrop}
                >
                  <span className="tool-button-content">
                    <span className="tool-btn-title">{tool.label}</span>
                    <span className="tool-btn-sub">{tool.hint}</span>
                  </span>
                </button>
              ))}
            </div>
          </div>

          <div className="tool-section">
            <label>{`\u0627\u0644\u0633\u0645\u0643:`}</label>
            <div className="size-buttons">
              {SIZE_OPTIONS.map((size) => (
                <button
                  key={size.id}
                  type="button"
                  className={brushSize === size.id ? "active" : ""}
                  onClick={() => handleSizeChange(size.id)}
                  aria-label={size.hint}
                  title={size.hint}
                  disabled={isLoading || saving || applyingCrop}
                >
                  <span className="tool-button-content">
                    <span className="tool-btn-title">{size.label}</span>
                    <span className="tool-btn-sub">{size.hint}</span>
                  </span>
                </button>
              ))}
            </div>
          </div>

          <div className="tool-section">
            <label>{`\u0627\u0644\u0644\u0648\u0646:`}</label>
            <div className="color-palette">
              {COLORS.map((color) => (
                <button
                  key={color}
                  type="button"
                  className={`color-swatch ${selectedColor === color ? "active" : ""}`}
                  style={{ backgroundColor: color }}
                  onClick={() => setSelectedColor(color)}
                  disabled={isLoading || saving || applyingCrop}
                />
              ))}
              <button
                type="button"
                className="color-picker-btn"
                onClick={() => setShowColorPicker((prev) => !prev)}
                disabled={isLoading || saving || applyingCrop}
              >
                +
              </button>
            </div>
            {showColorPicker ? (
              <div className="custom-color-picker">
                <input
                  type="color"
                  value={selectedColor}
                  onChange={(event) => setSelectedColor(event.target.value)}
                  disabled={isLoading || saving || applyingCrop}
                />
              </div>
            ) : null}
          </div>

          {activeTool === "text" ? (
            <div className="tool-section text-input-section">
              <label>{`\u0627\u0644\u0646\u0635:`}</label>
              <input
                type="text"
                className="text-draft-input"
                value={textDraft}
                onChange={(event) => setTextDraft(event.target.value)}
                placeholder={"\u0627\u0643\u062A\u0628\u064A \u0627\u0644\u0646\u0635 \u062B\u0645 \u0627\u0636\u063A\u0637\u064A \u0639\u0644\u0649 \u0627\u0644\u0635\u0648\u0631\u0629"}
                disabled={isLoading || saving || applyingCrop}
              />
            </div>
          ) : null}

          <div className="tool-section">
            <button type="button" className="action-btn" onClick={handleUndo} disabled={isLoading || saving || applyingCrop}>
              {`\u062A\u0631\u0627\u062C\u0639`}
            </button>
            <button type="button" className="action-btn" onClick={handleClear} disabled={isLoading || saving || applyingCrop}>
              {`\u0645\u0633\u062D`}
            </button>
          </div>
        </div>

        <div className="annotator-canvas-container">
          {isLoading ? <div className="loading-indicator">Loading...</div> : null}
          {loadError ? (
            <div className="error-indicator">
              Error: {loadError}
              <br />
              <button type="button" onClick={closeHandler} style={{ marginTop: "10px" }}>
                Close
              </button>
            </div>
          ) : null}
          <div
            className="annotator-stage"
            style={{
              width: displayDimensions.width || undefined,
              height: displayDimensions.height || undefined
            }}
          >
            {previewUrl && !loadError ? <img className="annotator-preview-image" src={previewUrl} alt="" draggable={false} /> : null}
            <canvas ref={canvasRef} />
          </div>
        </div>

        <div className="annotator-footer">
          <button
            type="button"
            className="cancel-btn"
            onClick={closeHandler}
            disabled={isLoading || saving || disabled || applyingCrop}
          >
            Cancel
          </button>
          <button
            type="button"
            className="save-btn"
            onClick={handleSave}
            disabled={isLoading || Boolean(loadError) || saving || disabled || applyingCrop}
          >
            {saving ? "Saving..." : "Save"}
          </button>
        </div>
      </div>
    </div>
  );
}
