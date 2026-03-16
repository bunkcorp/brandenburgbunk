// Dreamy 1-minute render of friction_string_curvature_ensemble.scd.
(
var patchPath;
var outPath;
var recSecs = 60;

patchPath = "/Users/kevinwoods/Desktop/ActuarialExams/Audio/supercollider/friction_string_curvature_ensemble.scd";
outPath = "/Users/kevinwoods/Desktop/friction_string_curvature_ensemble_1min_dreamy.wav";

s.options.outDevice = "Kevin’s AirPods P";
s.options.numInputBusChannels = 0;
s.options.blockSize = 512;

s.waitForBoot {
    var diskBuf;
    var recSynth;

    SynthDef(\diskRecorder, { |bufnum|
        DiskOut.ar(bufnum, In.ar(0, 2));
    }).add;
    s.sync;

    thisProcess.interpreter.executeFile(patchPath);

    Routine({
        // Let ensemble initialize, then retune to light/ethereal profile.
        1.0.wait;
        ~baseScale = [60, 62, 64, 67, 69, 71, 74, 76];
        if(~ensembleRoutine.notNil) { ~ensembleRoutine.stop; };
        ~ensembleRoutine = Routine({
            var t = 0;
            var n;
            loop {
                ~voices.do { |v, i|
                    n = ~baseScale.wrapAt(i + t) + [-12, 0, 7, 12].wchoose([0.08, 0.62, 0.20, 0.10]);
                    n = n.clip(55, 86);
                    v.set(
                        \freq, (n + rrand(-0.04, 0.04)).midicps,
                        \pressure, rrand(0.42, 0.62),
                        \bowRate, rrand(0.03, 0.11),
                        \bodyTilt, exprand(0.95, 1.08),
                        \amp, 0.07 + (0.03.rand)
                    );
                };
                t = t + 1;
                rrand(3.8, 7.2).wait;
            }
        }).play(AppClock);
    }).play(AppClock);

    diskBuf = Buffer.alloc(s, 131072, 2);
    s.sync;
    diskBuf.write(outPath, "wav", "int16", 0, 0, true);
    s.sync;

    recSynth = Synth.tail(s, \diskRecorder, [\bufnum, diskBuf.bufnum]);
    ("record_friction_string_ensemble_1min_dreamy: recording -> " ++ outPath).postln;

    Routine({
        recSecs.wait;
        recSynth.free;
        diskBuf.close;
        diskBuf.free;
        if(~ensembleRoutine.notNil) { ~ensembleRoutine.stop; };
        if(~voices.notNil) { ~voices.do(_.free); };
        "record_friction_string_ensemble_1min_dreamy: done".postln;
        s.quit;
        1.wait;
        0.exit;
    }).play(AppClock);
};
)
