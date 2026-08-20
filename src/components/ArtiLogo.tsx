import Image from "next/image";
import Link from "next/link";
import {
  ARTI_LOGO,
  ARTI_LOGO_DISPLAY_HEIGHT,
  artiLogoDisplayWidth,
} from "@/lib/arti-logo";

type Props = {
  showSubtitle?: boolean;
  className?: string;
  linked?: boolean;
};

export function ArtiLogo({ showSubtitle = true, className = "", linked = true }: Props) {
  const logoWidth = artiLogoDisplayWidth();

  const content = (
    <span className={`inline-flex items-center gap-3 ${className}`}>
      <Image
        src={ARTI_LOGO}
        alt="ARTI Mühendislik"
        width={logoWidth}
        height={ARTI_LOGO_DISPLAY_HEIGHT}
        className="h-8 w-auto shrink-0 object-contain"
        priority
      />
      {showSubtitle && (
        <span className="hidden text-lg font-semibold tracking-tight sm:inline">
          <span className="text-[var(--accent)]">Gantt</span>
        </span>
      )}
    </span>
  );

  if (!linked) return content;

  return (
    <Link
      href="/"
      className="rounded focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[var(--accent)]"
    >
      {content}
    </Link>
  );
}
