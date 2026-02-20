import { useEffect, useMemo, useState } from "react";
import { motion } from "motion/react";
import "./TypewriterEffect.css";

function joinClasses(...classes) {
  return classes.filter(Boolean).join(" ");
}

export default function TypewriterEffect({
  words = [],
  className,
  cursorClassName,
  typeSpeed = 70,
  deleteSpeed = 45,
  holdDelay = 1200
}) {
  const characters = useMemo(() => {
    const output = [];

    words.forEach((word, wordIndex) => {
      String(word?.text || "")
        .split("")
        .forEach((char) => {
          output.push({
            char,
            className: word?.className || ""
          });
        });

      if (wordIndex < words.length - 1) {
        output.push({
          char: " ",
          className: ""
        });
      }
    });

    return output;
  }, [words]);

  const [visibleCount, setVisibleCount] = useState(0);
  const [isDeleting, setIsDeleting] = useState(false);

  useEffect(() => {
    if (!characters.length) return undefined;

    let timeoutId;

    if (!isDeleting && visibleCount < characters.length) {
      timeoutId = window.setTimeout(() => {
        setVisibleCount((count) => Math.min(count + 1, characters.length));
      }, typeSpeed);
    } else if (!isDeleting && visibleCount >= characters.length) {
      timeoutId = window.setTimeout(() => {
        setIsDeleting(true);
      }, holdDelay);
    } else if (isDeleting && visibleCount > 0) {
      timeoutId = window.setTimeout(() => {
        setVisibleCount((count) => Math.max(count - 1, 0));
      }, deleteSpeed);
    } else if (isDeleting && visibleCount === 0) {
      timeoutId = window.setTimeout(() => {
        setIsDeleting(false);
      }, holdDelay / 2);
    }

    return () => window.clearTimeout(timeoutId);
  }, [characters.length, deleteSpeed, holdDelay, isDeleting, typeSpeed, visibleCount]);

  return (
    <div className={joinClasses("typewriter-root", className)} dir="ltr">
      <div className="typewriter-text">
        {characters.slice(0, visibleCount).map((item, idx) => (
          <span key={`char-${idx}`} className={joinClasses("typewriter-char", item.className)}>
            {item.char === " " ? <span className="typewriter-space">&nbsp;</span> : item.char}
          </span>
        ))}
      </div>

      <motion.span
        aria-hidden="true"
        className={joinClasses("typewriter-cursor", cursorClassName)}
        animate={{ opacity: [1, 0, 1] }}
        transition={{ duration: 0.8, repeat: Infinity, ease: "easeInOut" }}
      />
    </div>
  );
}
