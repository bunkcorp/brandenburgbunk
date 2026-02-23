// Brandenburg Liquid DnB v3 — production ready
// Requires: Quarks.install("wslib") then recompile (Cmd+K). Stop: ~stopBrandenburgMIDI.()

s.options.numInputBusChannels = 0;
s.options.maxNodes = 8192;

s.waitForBoot {

    //////////////////////////////
    // FILES
    //////////////////////////////

    ~midiDir = "/Users/kevinwoods/Downloads/brand5-3-mids";
    ~partFiles = ["flute-score.mid", "cello-score.mid", "harpsichord-score.mid"];
    ~amenPath = "/Users/kevinwoods/Desktop/ActuarialExams/Audio/copied_patches/amen.wav";
    ~amenBeats = 4;           // beats in the loop (4 = 1 bar, 8 = 2 bars)
    ~vocalPath = "/Users/kevinwoods/Desktop/ActuarialExams/Exam_Prep/Exam PA/PA_Module_Materials/PA-Rap-Vocal.aiff";

    //////////////////////////////
    // BUSSES
    //////////////////////////////

    ~verbBus   = Bus.audio(s, 2);
    ~masterBus = Bus.audio(s, 2);
    ~drumBus   = Bus.audio(s, 2);
    ~triggerBus= Bus.audio(s, 1);
    ~levelBus  = Bus.control(s, 1);

    //////////////////////////////
    // REVERB + BUILD SWEEP
    //////////////////////////////

    SynthDef(\lushVerb, { |verbBus=0, masterBus=0, intro=2, build=8, hold=220|
        var dry, chorus, sweep, wet;
        dry = In.ar(verbBus, 2);
        chorus = dry + DelayC.ar(dry, 0.03, LFTri.kr(0.2).range(0.005, 0.018)) * 0.3;
        sweep = EnvGen.kr(Env([400,400,6000,6000], [intro,build,hold], \lin));
        chorus = LPF.ar(chorus, sweep);
        wet = FreeVerb2.ar(chorus[0], chorus[1], 0.5, 0.9, 0.3);
        Out.ar(masterBus, dry + wet * 0.6);
    }).add;

    //////////////////////////////
    // MULTIBAND SIDECHAIN (mids duck, sub stays stable)
    //////////////////////////////

    SynthDef(\sidechainTrigger, { |triggerBus=0, tempo=2|
        var trig = Impulse.kr(tempo);
        var env  = EnvGen.ar(Env.perc(0.001, 0.18), trig);
        Out.ar(triggerBus, env);
    }).add;

    SynthDef(\multibandSidechain, {
        |masterBus=0, drumBus=0, triggerBus=0, levelBus=0, duck=0.5|
        var sig, drums, sc, level, low, mid, gain;
        sig   = In.ar(masterBus, 2);
        drums = In.ar(drumBus, 2);
        sc    = In.ar(triggerBus, 1);
        level = In.kr(levelBus, 1);
        low   = LPF.ar(sig, 120);
        mid   = HPF.ar(sig, 120);
        gain  = 1 - (sc * duck);
        mid   = mid * gain;
        Out.ar(0, ((low + mid) + drums) * level);
    }).add;

    //////////////////////////////
    // ARRANGEMENT LEVEL
    //////////////////////////////

    SynthDef(\arrangementLevel, { |levelBus=0, intro=2, build=8, hold=220, breakdown=10|
        var env = Env([0,0,1,1,0], [intro,build,hold,breakdown], \sin);
        Out.kr(levelBus, EnvGen.kr(env, doneAction:2));
    }).add;

    //////////////////////////////
    // ORCHESTRA
    //////////////////////////////

    SynthDef(\brand_flute, { |freq=440, amp=0.3, gate=1, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env.adsr(0.05,0.3,0.7,0.4), gate, doneAction:2);
        sig = Saw.ar(freq * [1,1.003]) * 0.5;
        sig = LPF.ar(sig.sum, 6000);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    SynthDef(\brand_cello, { |freq=220, amp=0.4, gate=1, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env.adsr(0.1,0.4,0.8,0.5), gate, doneAction:2);
        sig = Saw.ar(freq * [1,1.01]).sum;
        sig = LPF.ar(sig, 1200);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    SynthDef(\brand_harpsichord, { |freq=440, amp=0.25, gate=1, verbBus=0|
        var env, sig, width;
        env = EnvGen.kr(Env.perc(0.003,0.3), doneAction:2);
        sig = HPF.ar(Pulse.ar(freq,0.5), 500);
        width = SinOsc.kr(0.05).range(0.3,0.8);
        Out.ar(verbBus, Splay.ar(sig * env * amp, width));
    }).add;

    //////////////////////////////
    // REESE + SUB
    //////////////////////////////

    SynthDef(\reeseBass, { |freq=55, amp=0.7, gate=1, masterBus=0|
        var env, detune, reese, sub, sig;
        env = EnvGen.kr(Env.adsr(0.01,0.3,0.8,0.2), gate, doneAction:2);
        detune = [0.99,1.01];
        reese  = Saw.ar(freq * detune).sum;
        reese = LPF.ar(reese, 800);
        reese = reese.tanh;
        sub = SinOsc.ar(freq, 0, 0.5);
        sig = LeakDC.ar(reese + sub);
        Out.ar(masterBus, Pan2.ar(sig * env * amp));
    }).add;

    //////////////////////////////
    // DRUMS (Amen only; no synthetic kick/snare)
    //////////////////////////////

    SynthDef(\amenBreak, { |drumBus=0, buf=0, rate=1, amp=0.8, barRate=0.71667|
        var trig = Impulse.kr(barRate);
        var sig = PlayBuf.ar(1, buf, rate, trigger: trig, startPos: 0) * amp;
        Out.ar(drumBus, Pan2.ar(sig));
    }).add;

    SynthDef(\vocalPlay, { |buf=0, levelBus=0, amp=0.6|
        var sig = PlayBuf.ar(1, buf, 1, doneAction: 2);
        var level = In.kr(levelBus, 1);
        Out.ar(0, Pan2.ar(sig * amp * level, 0));
    }).add;

    s.sync;

    TempoClock.default.tempo = 172/60;

    //////////////////////////////
    // FX CHAIN
    //////////////////////////////

    ~levelBus.set(0);

    ~verbSynth = Synth(\lushVerb, [
        \verbBus, ~verbBus.index,
        \masterBus, ~masterBus.index,
        \intro, 2, \build, 8, \hold, 220
    ], s, \addToTail);

    ~triggerSynth = Synth(\sidechainTrigger, [
        \triggerBus, ~triggerBus.index,
        \tempo, TempoClock.default.tempo
    ], s, \addToTail);

    ~sidechain = Synth(\multibandSidechain, [
        \masterBus, ~masterBus.index,
        \drumBus, ~drumBus.index,
        \triggerBus, ~triggerBus.index,
        \levelBus, ~levelBus.index,
        \duck, 0.35
    ], s, \addToTail);

    //////////////////////////////
    // MIDI LOAD
    //////////////////////////////

    ~patterns = [];

    ~partFiles.do { |name|
        var path = ~midiDir +/+ name;
        if(PathName(path).isFile){
            var midi = SimpleMIDIFile.read(path);
            ~patterns = ~patterns.add([name, midi.p]);
            ("Loaded: " ++ name).postln;
        };
    };

    ~makePart = { |name, pat|
        var synth = \brand_flute;
        if(name.contains("cello")) { synth = \brand_cello };
        if(name.contains("harpsichord")) { synth = \brand_harpsichord };

        Pbindf(pat,
            \instrument, synth,
            \dur, (Pkey(\dur) * 0.25).round(0.25),
            \sustain, Pkey(\dur) * 0.95,
            \amp, (Pkey(\amp) ? ((Pkey(\vel) ? 100) / 127)) * 1.15,
            \verbBus, ~verbBus.index,
            \hasGate, (synth != \brand_harpsichord)
        );
    };

    ~orchestra = Ppar(~patterns.collect{|p| ~makePart.(p[0],p[1])});

    ~bass = Pbind(
        \instrument, \reeseBass,
        \scale, Scale.minor,
        \degree, Pseq([0,-2,-3,-5], inf),
        \octave, 2,
        \dur, 0.5,
        \masterBus, ~masterBus.index
    );

    // Amen: sample-accurate rate from frames; retrigger every bar (quant + barRate)
    ~barSec = 4 * 60 / 172;
    ~amenLoopSec = ~amenBeats * 60 / 172;
    ~drums = Pbind(
        \instrument, \amenBreak,
        \dur, 9999,
        \drumBus, ~drumBus.index,
        \buf, Pfunc { ~amenBuf.bufnum },
        \rate, Pfunc {
            var targetFrames = ~amenLoopSec * s.sampleRate;
            ~amenBuf.numFrames / targetFrames
        },
        \barRate, 1 / ~barSec,
        \amp, 0.25
    );

    ~score = Ppar([~orchestra,~bass,~drums]);

    ~startBrandenburg = {
        ~player = ~score.play(quant: 4);
        Synth(\arrangementLevel, [\levelBus, ~levelBus.index], s, \addToTail);
        "Brandenburg Liquid DnB v3 running @ 172 BPM".postln;
        nil
    };

    ~stopBrandenburgMIDI = {
        ~player.stop;
        fork { 2.wait; s.freeAll; "Stopped.".postln };
        nil
    };

    // Recording: path, duration (sec), openWhenDone (run from terminal = auto-record + open)
    ~recordToWav = { |path, duration = 120, openWhenDone = false|
        path = path ?? { "~/Desktop/Brandenburg_LiquidDnB_v3.wav".standardizePath };
        path = path.asString.standardizePath;
        if(s.serverRunning.not) { "Server not running. Boot server first.".postln; ^nil };
        s.record(path);
        ("Recording to " ++ path).postln;
        fork({
            duration.wait;
            s.stopRecording;
            ("Saved: " ++ path).postln;
            if(openWhenDone) { ("open " ++ path.quote).unixCmd };
        }, SystemClock);
        ("Stop early: s.stopRecording").postln;
        nil
    };

    // Load Amen (+ optional PA rap vocal); when ready, start playback and record
    ~recordPath = "~/Desktop/MyBrandenburg.wav".standardizePath;
    ~recordSec = 250;
    ~vocalBuf = nil;
    ~needVocal = PathName(~vocalPath).isFile;
    ~buffersReady = 0;
    ~tryStart = {
        ~buffersReady = ~buffersReady + 1;
        if((~needVocal.not).or(~buffersReady >= 2), {
            ~startBrandenburg.();
            if(~vocalBuf.notNil, {
                Synth(\vocalPlay, [\buf, ~vocalBuf.bufnum, \levelBus, ~levelBus.index, \amp, 0.6], s, \addToTail);
                "PA rap vocal started.".postln;
            });
            ~recordToWav.(~recordPath, ~recordSec, true);
        });
    };
    ~amenBuf = Buffer.read(s, ~amenPath, action: { "Amen loaded.".postln; ~tryStart.() });
    if(~needVocal, {
        ~vocalBuf = Buffer.read(s, ~vocalPath, action: { "Vocal loaded.".postln; ~tryStart.() });
    });

    nil
};
// ~stopBrandenburgMIDI.();
// Manual record: ~recordToWav.();  or ~recordToWav.("~/Desktop/MyBrandenburg.wav", 250, true)
