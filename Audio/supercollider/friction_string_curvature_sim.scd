// Friction string / curvature-style ambient instrument study (improved).
// Device strategy: try AirPods first, then fall back to MacBook speakers.

(
~configureFrictionServer = { |outDevice|
    s.options.numInputBusChannels = 0;
    s.options.outDevice = outDevice;
    // Do not open an input device for this patch; avoids Bluetooth profile churn.
    if(outDevice.asString.contains("MacBook Pro Speakers")) {
        s.options.sampleRate = 44100;
    };
    s.options.blockSize = 512;
    ("friction_string_curvature_sim: trying output -> " ++ outDevice).postln;
};

~startFrictionPatch = {
    SynthDef(\frictionString, {
        |out = 0, freq = 160, pressure = 0.9, bowRate = 0.25, amp = 1.8|

        var bowNoise;
        var wheel;
        var excite;
        var string;
        var body;
        var drift;
        var drift2;
        var feedback;
        var reverb;
        var present;
        var freqSmooth;
        var scrape;
        var dryStereo;

        // Mechanical wobble (wheel rotation imperfection).
        drift = LFNoise2.kr(0.15).range(0.97, 1.03);
        drift2 = SinOsc.kr(0.04).range(0.98, 1.02);
        // Smooth control-rate note changes.
        freqSmooth = Lag.kr(freq, 1.8);

        // Rotating wheel contact profile.
        wheel = LFTri.kr(bowRate + (LFNoise1.kr(0.3) * 0.05)).range(0, 1);

        // Stick-slip style friction approximation.
        bowNoise = PinkNoise.ar * pressure;
        bowNoise = bowNoise * (wheel + 0.1);
        bowNoise = LeakDC.ar(bowNoise);

        // Nonlinear bow interaction.
        excite = tanh(bowNoise * 6);
        // Continuous low-energy injection keeps long drones alive.
        excite = excite + (SinOsc.ar(freqSmooth * 0.5) * 0.02);

        // String resonance.
        string = CombL.ar(
            excite,
            0.2,
            (1 / (freqSmooth * drift * drift2)).clip(0.001, 0.2),
            10
        );

        // Multi-mode resonant body (wood plate simulation).
        body =
            BPF.ar(string, freqSmooth * 1.0, 0.2) +
            BPF.ar(string, freqSmooth * 1.61, 0.2) +
            BPF.ar(string, freqSmooth * 2.32, 0.25) +
            BPF.ar(string, freqSmooth * 3.73, 0.3);
        // Add a little dry string back so the tone stays present.
        body = body + (LPF.ar(string, 3500) * 0.35);
        // Rosin-like scraping detail for realism.
        scrape = HPF.ar(PinkNoise.ar * pressure, 4000) * 0.05;
        body = body + scrape;

        // Subtle sustain feedback.
        feedback = LocalIn.ar(1);
        body = body + (feedback * 0.22);
        LocalOut.ar(Limiter.ar(LeakDC.ar(Mix(body)), 0.9));

        // Ambient diffusion (plugin-safe fallback to FreeVerb2).
        reverb = FreeVerb2.ar(body, body, 0.6, 0.92, 0.4);
        // Brightness layer helps small headphones cut through.
        present = HPF.ar(body, 1200) * 0.3;
        // Slight stereo spread for width.
        dryStereo = Splay.ar([body, DelayC.ar(body, 0.03, 0.011)]);

        Out.ar(out, Limiter.ar((dryStereo + reverb + present) * amp, 0.95));
    }).add;

    SynthDef(\granularCloud, {
        |out = 0, inChan = 0, trigRate = 6, grainDur = 0.25, amp = 0.5|

        var source;
        var grains;
        var spread;
        var reverb;

        source = SoundIn.ar(inChan);

        // Stereo motion.
        spread = LFNoise1.kr(0.2);

        grains = GrainIn.ar(
            2,
            Impulse.kr(trigRate),
            grainDur,
            source,
            spread
        );

        reverb = FreeVerb2.ar(grains[0], grains[1], 0.75, 0.92, 0.45);

        Out.ar(out, reverb * amp);
    }).add;

    s.sync;

    "friction_string_curvature_sim: server booted".postln;

    // Play friction string.
    x = Synth(\frictionString);
    "friction_string_curvature_sim: friction synth started".postln;
    // Audible route check ping (one short chirp), optional for renders.
    if((~frictionRoutePing ? true), {
        {
            var env = Env.perc(0.01, 0.45).ar(doneAction: 2);
            var sig = SinOsc.ar(990) * 0.25 * env;
            Out.ar(0, sig ! 2);
        }.play;
    });

    // Slow evolving pitch motion so it is not a single-note drone.
    ~noteSeq = [43, 45, 48, 50, 52, 55, 57, 60];
    if(~motionRoutine.notNil) { ~motionRoutine.stop; };
    ~motionRoutine = Routine({
        var idx = 0;
        var n;
        loop {
            n = ~noteSeq.wrapAt(idx) + [-12, 0, 12].wchoose([0.1, 0.75, 0.15]);
            n = n.clip(40, 72); // Keep melody in consistently audible register.
            ("friction_string_curvature_sim: new note -> " ++ n).postln;
            x.set(
                \freq, (n + rrand(-0.15, 0.15)).midicps,
                \pressure, rrand(0.72, 0.95),
                \bowRate, rrand(0.08, 0.32)
            );
            idx = idx + 1;
            rrand(3.5, 8.0).wait;
        }
    }).play(AppClock);

    // Optional controls:
    // x.set(\freq, 70);
    // x.set(\pressure, 0.9);
    // x.set(\bowRate, 0.1);

    // Optional granular layer:
    // y = Synth(\granularCloud);

    // Stop:
    // x.free;
    // ~motionRoutine.stop;
    // y.free;
};

~bootFrictionWithFallback = { |primaryOut, fallbackOut|
    ~configureFrictionServer.(primaryOut);
    s.waitForBoot(
        {
            ~startFrictionPatch.();
        },
        100,
        {
            ("friction_string_curvature_sim: failed on -> " ++ primaryOut).warn;
            if(fallbackOut.notNil) {
                AppClock.sched(0.5, {
                    ~bootFrictionWithFallback.(fallbackOut, nil);
                    nil
                });
            };
        }
    );
};

~bootFrictionWithFallback.("Kevin’s AirPods P", "MacBook Pro Speakers");
)
