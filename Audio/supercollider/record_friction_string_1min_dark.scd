// Light/ethereal 1-minute render of friction_string_curvature_sim.scd.
(
var patchPath;
var outPath;
var recSecs = 60;

patchPath = "/Users/kevinwoods/Desktop/ActuarialExams/Audio/supercollider/friction_string_curvature_sim.scd";
outPath = "/Users/kevinwoods/Desktop/friction_string_curvature_1min_dreamy.wav";

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

    // Disable startup ping in render.
    ~frictionRoutePing = false;
    thisProcess.interpreter.executeFile(patchPath);

    // Brighter sequence and airy bow motion.
    ~noteSeq = [60, 62, 64, 67, 69, 71, 74, 76];
    if(~motionRoutine.notNil) { ~motionRoutine.stop; };
    ~motionRoutine = Routine({
        var idx = 0;
        var n;
        loop {
            n = ~noteSeq.wrapAt(idx) + [-12, 0, 7, 12].wchoose([0.1, 0.6, 0.2, 0.1]);
            n = n.clip(55, 84);
            x.set(
                \freq, (n + rrand(-0.05, 0.05)).midicps,
                \pressure, rrand(0.45, 0.68),
                \bowRate, rrand(0.04, 0.12),
                \amp, 0.82
            );
            idx = idx + 1;
            rrand(4.0, 7.5).wait;
        }
    }).play(AppClock);

    diskBuf = Buffer.alloc(s, 131072, 2);
    s.sync;
    diskBuf.write(outPath, "wav", "int16", 0, 0, true);
    s.sync;

    recSynth = Synth.tail(s, \diskRecorder, [\bufnum, diskBuf.bufnum]);
    ("record_friction_string_1min_dreamy: recording -> " ++ outPath).postln;

    Routine({
        recSecs.wait;
        recSynth.free;
        diskBuf.close;
        diskBuf.free;
        if(~motionRoutine.notNil) { ~motionRoutine.stop; };
        if(x.notNil) { x.free; };
        "record_friction_string_1min_dreamy: done".postln;
        s.quit;
        1.wait;
        0.exit;
    }).play(AppClock);
};
)
