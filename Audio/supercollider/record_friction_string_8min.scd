// Records friction_string_curvature_sim.scd for 1 minute to Desktop via DiskOut.
(
var patchPath;
var outPath;
var recSecs = 60;

patchPath = "/Users/kevinwoods/Desktop/ActuarialExams/Audio/supercollider/friction_string_curvature_sim.scd";
outPath = "/Users/kevinwoods/Desktop/friction_string_curvature_1min.wav";

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

    // Disable startup ping in the render.
    ~frictionRoutePing = false;
    thisProcess.interpreter.executeFile(patchPath);

    diskBuf = Buffer.alloc(s, 131072, 2);
    s.sync;
    diskBuf.write(outPath, "wav", "int16", 0, 0, true);
    s.sync;

    recSynth = Synth.tail(s, \diskRecorder, [\bufnum, diskBuf.bufnum]);
    ("record_friction_string_1min: recording -> " ++ outPath).postln;

    Routine({
        recSecs.wait;
        recSynth.free;
        diskBuf.close;
        diskBuf.free;
        if(~motionRoutine.notNil) { ~motionRoutine.stop; };
        if(x.notNil) { x.free; };
        "record_friction_string_1min: done".postln;
        s.quit;
        1.wait;
        0.exit;
    }).play(AppClock);
};
)
