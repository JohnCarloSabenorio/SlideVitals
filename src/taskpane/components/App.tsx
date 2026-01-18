import * as React from "react";

import { useState } from "react";
import Header from "./Header";
import Results from "./Results";
import { makeStyles } from "@fluentui/react-components";
import {
  Ribbon24Regular,
  LockOpen24Regular,
  DesignIdeas24Regular,
  TextSortAscending16Filled,
} from "@fluentui/react-icons";
interface AppProps {
  title: string;
}

const App: React.FC<AppProps> = (props: AppProps) => {
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.

  const [presentationErrors, setPresentationErrors] = useState({});
  const [globalFontNameError, setGlobalFontNameError] = useState("");
  const [globalFontSizeError, setGlobalFontSizeError] = useState("");
  const [isScanning, setIsScanning] = useState(false);

  async function scanSlideVitals() {
    await PowerPoint.run(async (context) => {
      setIsScanning(true);
      const slideCount = context.presentation.slides.getCount();
      await context.sync();

      const slideErrors = {};

      const globalFontSizes: Set<number> = new Set();
      const globalFontNames = new Set();
      // Check text shapes
      for (let i = 0; i < slideCount.value; i++) {
        const slide = context.presentation.slides.getItemAt(i);
        const shapes = slide.shapes;
        shapes.load("type");
        await context.sync();

        // load text frames
        shapes.load("items/textFrame/hasText");
        await context.sync();

        // filter text frames
        const textShapes = shapes.items.filter(
          (shape) => shape.textFrame && shape.textFrame.hasText
        );

        // load text ranges
        textShapes.forEach((textShape) => {
          textShape.textFrame.textRange.load("text");
          textShape.textFrame.textRange.load("font");
        });

        await context.sync();
        checkTexts(textShapes, i + 1);
      }

      // Check text Density

      setPresentationErrors(slideErrors);
      setIsScanning(false);

      function checkTexts(textShapes, slideNumber) {
        let charCount = 0;
        let usedFonts = [];
        let usedFontSizes = [];
        let fontNameErrors = [];
        let fontSizeWarnings = [];
        let fontSizeErrors = [];
        // Get total text count

        textShapes.forEach((textShape) => {
          const text = textShape.textFrame.textRange.text;

          // Check font name consistency
          globalFontNames.add(textShape.textFrame.textRange.font.name);
          usedFonts.push(textShape.textFrame.textRange.font.name);
          // Check font size consistency
          globalFontSizes.add(textShape.textFrame.textRange.font.size);
          usedFontSizes.push(textShape.textFrame.textRange.font.size);
          charCount += text.length;
        });

        // Check if presentation has more than 2 fonts
        if (globalFontNames.size > 2) {
          setGlobalFontNameError(
            `Inconsistent font. Slide should only use two main fonts for consistency. Recorded fonts: ${Array.from(globalFontNames)}`
          );
        }

        console.log("global font sizes:", globalFontSizes);
        if (globalFontSizes.size > 4) {
          setGlobalFontSizeError(
            `You should only use up to 4 font sizes for your presentation to maintain consistency. Recorded font sizes: ${Array.from(globalFontSizes)}`
          );
        }

        // Check if the values are within the range of minimum and maximum font size
        for (const fontSize of Array.from(globalFontSizes)) {
          if (fontSize < 12) {
            fontSizeErrors.push(`Your font size should be a minimum of 12.`);
            break;
          } else if (fontSize > 72) {
            fontSizeErrors.push(`Your font size should be a maximum of 72.`);
            break;
          }
        }

        let textDensityWarnings = [];
        let textDensityErrors = [];

        // Check the character count of a slide
        if (charCount > 400 && charCount <= 700) {
          textDensityWarnings.push(
            "Slide has too much text (400+ characters). Consider shortening or splitting the content."
          );
        } else if (charCount > 700) {
          textDensityErrors.push(
            "Slide is overloaded with text (700+ characters). Break this into multiple slides or move details to speaker notes."
          );
        }

        slideErrors[slideNumber] = {
          ...slideErrors[slideNumber],
          textDensityErrors,
          textDensityWarnings,
          fontNameErrors,
          fontSizeErrors,
        };
      }
    });
  }

  return (
    <div className="w-full h-[100vh] flex flex-col">
      <Header logo="assets/slidevitals-logo.png" title={props.title} />

      <button
        onClick={scanSlideVitals}
        className={`${isScanning ? "bg-gray-500" : "bg-orange-500 cursor-pointer"} text-white text-xl font-bold rounded-md px-5 py-3 mt-3`}
      >
        {isScanning ? "Scanning..." : "Scan Now"}
      </button>

      <Results
        presentationErrors={presentationErrors}
        globalFontNameError={globalFontNameError}
        globalFontSizeError={globalFontSizeError}
      />
    </div>
  );
};

export default App;
