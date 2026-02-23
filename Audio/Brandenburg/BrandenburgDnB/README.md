# Brandenburg DnB – SP1200 French Touch Engine

Classical sample → SP1200/S-760-style processing → DnB breaks → ethereal reverb/delay chain. Controllable via SuperCollider (OSC) or TidalCycles.

## Structure

| Path | Contents |
|------|----------|
| `engine/` | Sample loader (slice by bars/quarters), slice playback SynthDef |
| `dsp/` | SP-1200 emulation (12-bit, 26k, drive, noise), S-760 clean mode (lowpass, chorus) |
| `patterns/` | DnB kick/snare/hat SynthDefs, 172 BPM clock, pattern runner |
| `fx/` | Highpass → hall reverb → ping-pong delay → chorus → sidechain (duck to kick) |
| `main.scd` | Busses, chain wiring, OSC controls, demo |
| `run_brandenburg.scd` | Boot entry: load main, optional sample path + demo |

## Quick start

1. **SuperCollider**: Boot server, then run `run_brandenburg.scd` (e.g. open file, Cmd-Enter).
2. **Optional**: Set `samplePath` in `run_brandenburg.scd` to a WAV/AIFF (e.g. Brandenburg excerpt); re-run to load and start the demo.
3. **Manual**:  
   `~loadBrandenburg.("/path/to.wav", 172, 16)`  
   `~playSlice.(0, 1, 0, 1)`  
   `~runDnBDemo.()`  
   `~stopDnB.()`

## OSC (live / Tidal / Bela)

| Address | Args | Meaning |
|---------|------|---------|
| `/brandenburg/dsp_mode` | 0 or 1 | 0 = S-760 clean, 1 = SP1200 |
| `/brandenburg/play_slice` | index, speed, reverse, duration | Play slice to engine |
| `/brandenburg/fx/rev_decay` | float | Reverb decay (e.g. 4.5–7.5) |
| `/brandenburg/fx/rev_wet` | float | Reverb wet (e.g. 0.2–0.35) |
| `/brandenburg/fx/delay_fb` | float | Delay feedback (e.g. 0.35–0.45) |
| `/brandenburg/filter_cutoff` | float | S-760 lowpass (Hz) |
| `/brandenburg/sp_drive` | float | SP1200 drive (0–1) |

## FX chain

Break/sample → **Highpass 120 Hz** → **Hall reverb** (decay 5.5 s, wet ~28%, highcut 6 kHz) → **Ping-pong delay** (1/8 dotted @ 172 BPM, feedback 40%, wet 20%) → **Chorus** → **Sidechain** (duck to kick). Reverb bus is sidechained to kick for a tighter mix.

## Dependencies

- SuperCollider 3.10+ (GVerb, Buffer, etc.).
- No extra quarks required for the core engine; TidalCycles is optional for pattern control.

## Optional (Kevin-level)

- Granular freeze, probabilistic break mutations, Markov slice rearrangement, filter resonance tied to kick can be added as extra SynthDefs and OSC params in `dsp/` and `main.scd`.
