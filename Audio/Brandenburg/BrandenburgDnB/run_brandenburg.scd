/*
  Boot script for Brandenburg DnB engine.
  Start SuperCollider, boot server, then run this file (e.g. Cmd-Enter on "run_brandenburg.scd").
  Optionally set a sample path below for the demo.
*/
(
var projectDir = PathName(thisProcess.nowExecutingPath).pathOnly;
var samplePath = ""; // e.g. "/path/to/brandenburg_excerpt.wav"

s.waitForBoot {
  (projectDir +/+ "main.scd").load;
  s.sync;
  // Optional: load sample and run full demo
  if (samplePath != "") {
    ~loadBrandenburg.(samplePath, 172, 16);
    fork {
      2.wait;
      ~playSlice.(0, 1, 0, 2);
      ~runDnBDemo.();
      "Demo running: Brandenburg slice + DnB break + ethereal FX. Stop: ~stopDnB.().".postln;
    };
  } {
    "Brandenburg DnB loaded. Set samplePath in run_brandenburg.scd and re-run, or: ~loadBrandenburg(\"/path/to.wav\"); ~playSlice.(0,1,0,1); ~runDnBDemo.().".postln;
  };
};
)
