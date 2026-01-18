import React from "react";
export default function Slide({
  slideNumber,
  slideProblems,
}: {
  slideNumber: number;
  slideProblems: any;
}) {
  return (
    <div className="bg-white p-3">
      <h1 className="font-bold">Slide {slideNumber}</h1>

      {(slideProblems.textDensityWarnings.length != 0 ||
        slideProblems.textDensityErrors.length != 0) && (
        <h1 className="text-3xl mt-3">Text Density</h1>
      )}

      {/* Text Density Warnings */}
      <div className="mt-3">
        <ul className="list-disc text-amber-400 font-semibold mx-3">
          {slideProblems.textDensityWarnings &&
            slideProblems.textDensityWarnings.map((error, idx) => {
              return <li key={idx}>{error}</li>;
            })}
        </ul>
      </div>

      {/* Text Density Errors */}
      <div className="mt-3">
        <ul className="list-disc text-red-500 font-semibold mx-3">
          {slideProblems.textDensityErrors &&
            slideProblems.textDensityErrors.map((error, idx) => {
              return <li key={idx}>{error}</li>;
            })}
        </ul>
      </div>

      {slideProblems.fontNameErrors.length != 0 && <h1 className="text-3xl mt-3">Font Family</h1>}

      {/* <div className="mt-3">
        <ul className="list-disc text-red-500 font-semibold mx-3">
          {slideProblems.fontNameErrors.map((error, idx) => {
            return <li key={idx}>{error}</li>;
          })}
        </ul>
      </div> */}
    </div>
  );
}
