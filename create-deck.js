const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const path = require("path");
const fs = require("fs");

// Icon imports
const { FaRocket, FaGlobeAmericas, FaUserAstronaut, FaBolt, FaUsers, FaGraduationCap, FaSpaceShuttle, FaStar, FaHandshake } = require("react-icons/fa");
const { GiLightsaber, GiArtificialIntelligence } = require("react-icons/gi"); // fallback if available, else use bolt/star
const { MdOutlineScience, MdPrecisionManufacturing } = require("react-icons/md");

// Color palette (no # prefix)
const COLORS = {
  bgDark: "0B1120",
  bgCard: "1E2937",
  gold: "FBBF24",
  flame: "F97316",
  sky: "38BDF8",
  white: "F1F5F9",
  muted: "94A3B8",
  navy: "1E3A5F",
  textDark: "0F172A"
};

async function iconToBase64Png(IconComponent, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

async function createPresentation() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Starfleet Padawan Academy Concept";
  pres.title = "Starfleet Padawan Academy at SpaceX";
  pres.subject = "Vision Pitch Deck";

  const projectRoot = __dirname;
  const imgDir = path.join(projectRoot, "images");

  // Pre-render icons
  const iconRocket = await iconToBase64Png(FaRocket, "#FBBF24", 256);
  const iconGlobe = await iconToBase64Png(FaGlobeAmericas, "#38BDF8", 256);
  const iconUser = await iconToBase64Png(FaUserAstronaut, "#F97316", 256);
  const iconBolt = await iconToBase64Png(FaBolt, "#FBBF24", 256);
  const iconUsers = await iconToBase64Png(FaUsers, "#38BDF8", 256);
  const iconGrad = await iconToBase64Png(FaGraduationCap, "#FBBF24", 256);
  const iconStar = await iconToBase64Png(FaStar, "#FBBF24", 256);
  const iconHandshake = await iconToBase64Png(FaHandshake, "#38BDF8", 256);

  // ========== SLIDE 1: TITLE ==========
  let slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  // Logo
  slide.addImage({
    path: path.join(imgDir, "logo.jpg"),
    x: 3.5, y: 0.3, w: 3, h: 3,
    sizing: { type: "contain", w: 3, h: 3 }
  });

  slide.addText("STARFLEET\nPADAWAN ACADEMY", {
    x: 0.5, y: 3.4, w: 9, h: 1.4,
    fontSize: 42, fontFace: "Arial", bold: true, color: COLORS.white,
    align: "center", valign: "middle", margin: 0
  });

  slide.addText("AT SPACEX", {
    x: 0.5, y: 4.7, w: 9, h: 0.5,
    fontSize: 24, fontFace: "Arial", color: COLORS.gold,
    align: "center", bold: true
  });

  // ========== SLIDE 2: THE VISION (Musk Quote) ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addText("THE VISION", {
    x: 0.5, y: 0.4, w: 9, h: 0.6,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true,
    charSpacing: 4
  });

  slide.addText([
    { text: '"The purpose of SpaceX is that we want to make Star Trek real.', options: { breakLine: true } },
    { text: 'We want to make Starfleet Academy real, so that it\'s not always science fiction."', options: { breakLine: true } }
  ], {
    x: 0.8, y: 1.3, w: 8.4, h: 2.2,
    fontSize: 28, fontFace: "Georgia", color: COLORS.white, italic: true,
    align: "left"
  });

  slide.addText("— Elon Musk", {
    x: 0.8, y: 3.7, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: "Arial", color: COLORS.gold, align: "left"
  });

  slide.addText("Eureka realized. The ships are coming. The people who will fly them must be forged now.", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.8,
    fontSize: 16, fontFace: "Arial", color: COLORS.muted
  });

  // ========== SLIDE 3: THE SYNTHESIS ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addText("THE SYNTHESIS", {
    x: 0.5, y: 0.3, w: 9, h: 0.5,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 4
  });

  slide.addText("Three worlds. One destiny.", {
    x: 0.5, y: 0.8, w: 9, h: 0.6,
    fontSize: 32, fontFace: "Arial", bold: true, color: COLORS.white
  });

  const pillars = [
    { title: "STARFLEET", desc: "Exploration. Command. Science. The Prime Directive. Boldly going where the math says we can — and the heart says we must.", icon: iconGlobe, color: COLORS.sky },
    { title: "PADAWAN", desc: "Apprenticeship. Precision. Resilience. The Living Force. Inner discipline that survives when the hull breaches and comms go dark.", icon: iconBolt, color: COLORS.gold },
    { title: "SPACEX", desc: "First principles. Rapid iteration. Hardware reality. Starship-scale ambition. Making life multi-planetary in our lifetime.", icon: iconRocket, color: COLORS.flame }
  ];

  pillars.forEach((p, i) => {
    const x = 0.5 + (i * 3.1);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.7, w: 2.9, h: 3.5,
      fill: { color: COLORS.bgCard },
      rectRadius: 0.1,
      shadow: { type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.25 }
    });
    slide.addImage({ data: p.icon, x: x + 1.05, y: 1.95, w: 0.7, h: 0.7 });
    slide.addText(p.title, {
      x: x + 0.15, y: 2.8, w: 2.6, h: 0.5,
      fontSize: 16, fontFace: "Arial", bold: true, color: p.color, align: "center"
    });
    slide.addText(p.desc, {
      x: x + 0.15, y: 3.35, w: 2.6, h: 1.7,
      fontSize: 12, fontFace: "Arial", color: COLORS.white, align: "left", valign: "top"
    });
  });

  // ========== SLIDE 4: CAMPUS (image slide) ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({
    path: path.join(imgDir, "campus.jpg"),
    x: 0, y: 0, w: 10, h: 5.625,
    sizing: { type: "cover", w: 10, h: 5.625 }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.2, w: 10, h: 1.425,
    fill: { color: COLORS.bgDark, transparency: 40 }
  });

  slide.addText("THE CAMPUS", {
    x: 0.5, y: 4.35, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 3
  });
  slide.addText("Starbase, Texas — Where the temple meets the launch pad", {
    x: 0.5, y: 4.75, w: 9, h: 0.6,
    fontSize: 22, fontFace: "Arial", bold: true, color: COLORS.white
  });

  // ========== SLIDE 5: THE PROGRAM ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addText("THE PROGRAM", {
    x: 0.5, y: 0.3, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 4
  });
  slide.addText("4 Years Core + Padawan Apprenticeship", {
    x: 0.5, y: 0.7, w: 9, h: 0.6,
    fontSize: 28, fontFace: "Arial", bold: true, color: COLORS.white
  });

  const years = [
    { year: "YEAR 1", title: "Awakening", items: "All pillars foundation. The Forge. First 72-hour Lost Ship sim." },
    { year: "YEAR 2", title: "Refinement", items: "Specialization tracks. Real mission contributions. Cross-training." },
    { year: "YEAR 3", title: "Command", items: "Lead teams. Multi-week habitat mission. Select Master for apprenticeship." },
    { year: "YEAR 4", title: "Synthesis", items: "Capstone project. Public defense before mixed Council board." },
    { year: "PADAWAN", title: "Apprenticeship", items: "1-2 years embedded. Starship crew, Starbase ops, forward missions. Knighting." }
  ];

  years.forEach((y, i) => {
    const yPos = 1.5 + (i * 0.78);
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0.5, y: yPos, w: 0.08, h: 0.65,
      fill: { color: i === 4 ? COLORS.gold : COLORS.sky }
    });
    slide.addText(y.year, {
      x: 0.75, y: yPos, w: 1.5, h: 0.65,
      fontSize: 13, fontFace: "Arial", bold: true, color: COLORS.gold, valign: "middle"
    });
    slide.addText(y.title, {
      x: 2.3, y: yPos, w: 2.2, h: 0.65,
      fontSize: 15, fontFace: "Arial", bold: true, color: COLORS.white, valign: "middle"
    });
    slide.addText(y.items, {
      x: 4.6, y: yPos, w: 5, h: 0.65,
      fontSize: 13, fontFace: "Arial", color: COLORS.muted, valign: "middle"
    });
  });

  // ========== SLIDE 6: TRAINING (image) ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({
    path: path.join(imgDir, "training.jpg"),
    x: 0, y: 0, w: 10, h: 5.625,
    sizing: { type: "cover", w: 10, h: 5.625 }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.0, w: 10, h: 1.625,
    fill: { color: COLORS.bgDark, transparency: 35 }
  });

  slide.addText("THE FORGE + THE PATH", {
    x: 0.5, y: 4.15, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 3
  });
  slide.addText("Precision engineering meets lightsaber discipline — on the actual launch pad", {
    x: 0.5, y: 4.55, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Arial", bold: true, color: COLORS.white
  });

  // ========== SLIDE 7: MENTORSHIP ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({
    path: path.join(imgDir, "mentor.jpg"),
    x: 4.5, y: 0, w: 5.5, h: 5.625,
    sizing: { type: "cover", w: 5.5, h: 5.625 }
  });

  slide.addText("THE MASTER & THE PADAWAN", {
    x: 0.4, y: 0.4, w: 4, h: 0.5,
    fontSize: 12, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 3
  });
  slide.addText("Apprenticeship is the soul of the Academy.", {
    x: 0.4, y: 0.9, w: 4, h: 0.8,
    fontSize: 22, fontFace: "Arial", bold: true, color: COLORS.white
  });

  const mentorPoints = [
    "Every cadet is paired with a living Master during apprenticeship",
    "Masters are active operators: chief engineers, flight directors, Starship captains",
    "The relationship is lifelong — the chain of knowledge and character",
    "Graduation is not a degree. It is a knighting."
  ];

  mentorPoints.forEach((pt, i) => {
    slide.addImage({ data: iconStar, x: 0.4, y: 1.9 + (i * 0.75), w: 0.35, h: 0.35 });
    slide.addText(pt, {
      x: 0.9, y: 1.85 + (i * 0.75), w: 3.4, h: 0.7,
      fontSize: 13, fontFace: "Arial", color: COLORS.white, valign: "top"
    });
  });

  // ========== SLIDE 8: GRADUATES ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({
    path: path.join(imgDir, "graduates.jpg"),
    x: 0, y: 0, w: 10, h: 5.625,
    sizing: { type: "cover", w: 10, h: 5.625 }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.0, w: 10, h: 1.625,
    fill: { color: COLORS.bgDark, transparency: 35 }
  });

  slide.addText("THE FIRST GRADUATES", {
    x: 0.5, y: 4.15, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 3
  });
  slide.addText("Ensign-Knights ready to crew the fleet that makes the future real", {
    x: 0.5, y: 4.55, w: 9, h: 0.9,
    fontSize: 20, fontFace: "Arial", bold: true, color: COLORS.white
  });

  // ========== SLIDE 9: THE ASCENT (symbolic) ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({
    path: path.join(imgDir, "ascent.jpg"),
    x: 0, y: 0, w: 10, h: 5.625,
    sizing: { type: "cover", w: 10, h: 5.625 }
  });

  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 4.0, w: 10, h: 1.625,
    fill: { color: COLORS.bgDark, transparency: 30 }
  });

  slide.addText("THE LONG ASCENT", {
    x: 0.5, y: 4.15, w: 9, h: 0.4,
    fontSize: 14, fontFace: "Arial", color: COLORS.gold, bold: true, charSpacing: 3
  });
  slide.addText("The ships will come. The question is whether we will be ready — in mind, hand, and spirit.", {
    x: 0.5, y: 4.55, w: 9, h: 0.9,
    fontSize: 18, fontFace: "Georgia", italic: true, color: COLORS.white
  });

  // ========== SLIDE 10: CALL TO ACTION ==========
  slide = pres.addSlide();
  slide.background = { color: COLORS.bgDark };

  slide.addImage({ data: iconGrad, x: 4.5, y: 0.6, w: 1, h: 1 });

  slide.addText("BUILD THE ACADEMY", {
    x: 0.5, y: 1.9, w: 9, h: 0.7,
    fontSize: 36, fontFace: "Arial", bold: true, color: COLORS.white, align: "center"
  });

  slide.addText("This is more than a concept. It is the logical next step in making Star Trek real.", {
    x: 1, y: 2.7, w: 8, h: 0.6,
    fontSize: 16, fontFace: "Arial", color: COLORS.muted, align: "center"
  });

  const ctas = [
    "Recruit the first cohort of Padawans who have already built something real",
    "Stand up the Forge at Starbase alongside existing training infrastructure",
    "Identify the first Masters willing to take on apprentices for the long voyage",
    "Begin the conversation that turns vision into policy, budget, and flight manifest"
  ];

  ctas.forEach((c, i) => {
    slide.addText("→  " + c, {
      x: 1.2, y: 3.5 + (i * 0.42), w: 7.6, h: 0.4,
      fontSize: 14, fontFace: "Arial", color: COLORS.white
    });
  });

  slide.addText("May the Force be with the boosters — and the people who fly them.", {
    x: 0.5, y: 5.2, w: 9, h: 0.35,
    fontSize: 13, fontFace: "Georgia", italic: true, color: COLORS.gold, align: "center"
  });

  // Write the file
  const outPath = path.join(projectRoot, "Starfleet_Padawan_Academy_SpaceX.pptx");
  await pres.writeFile({ fileName: outPath });
  console.log("Presentation written to:", outPath);
}

createPresentation().catch(console.error);