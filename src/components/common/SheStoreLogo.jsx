import sheStoreLogoSrc from "../../assets/She Store Logo - Upscaled.png";

export default function SheStoreLogo({
  href = "#/orders",
  className = "",
  imageClassName = "",
  alt = "She-Store Logo",
  ariaLabel = "She-Store - شي ستور"
}) {
  const image = (
    <img
      src={sheStoreLogoSrc}
      alt={alt}
      className={imageClassName}
      loading="eager"
      decoding="async"
    />
  );

  if (!href) {
    return <span className={className}>{image}</span>;
  }

  return (
    <a href={href} className={className} aria-label={ariaLabel}>
      {image}
    </a>
  );
}
