(
s.waitForBoot({
	Routine{
		~serverLatency = 0.1;

		SynthDef("sampler",{
			arg buf,
			volumeDB=0,
			rate = 1.0,
			pitch = 0.0,
			gate = 1,
			t_trig = 1,
			xfade=0.01,
			bpmSource=120,
			bpmTarget=120,
			// retriggering
			retrigNumTotal=0,
			retrigRateChangeBeats=1,
			retrigRateStart=1,
			retrigRateEnd=0,
			retrigPitchChange=0,
			retrigVolumeChange=0,
			sliceAttackBeats = 0.001,
			sliceDurationBeats = 0,
			sliceReleaseBeats = 0.001,
			sliceNum = 0,
			sliceCount = 32, // number of slices to cut the sample into
			effectDry = 1.0,
			effectDryOut,
			effectComb = 0.0,
			effectCombOut,
			effectDelay = 0.0,
			effectReverse = 0,
			effectTimestretchStart = 0,
			effectTimestretchEnd = 0,
			effectTimestretchBeats = 0,
			effectReverb = 0.0,
			effectLPFStart = 20000,
			effectLPFEnd = 0,
			effectLPFBeats = 0,
			effectReverbOut
			;
			var snd,sndA,sndB,crossfade,aOrB;
			var sndDry, sndReverb, sndComb;
			var posA,posB;
			var retrigTrig, retrigCount;
			var frames=BufFrames.ir(buf);
			var seconds=BufDur.ir(buf)*bpmSource/bpmTarget;
			var sliceSeconds = seconds / sliceCount;
			var pos = 0.0;
			var sliceTrigger = t_trig;
			var beatDuration = 60 / bpmTarget;
			var retrigCountFeedback = 0;
			var timestretchPos, timestretchRate, effectTimestretch;

			sliceNum = sliceNum.mod(sliceCount);
			// convert slice beat to seconds
			pos = Select.kr(effectReverse>0,[
				sliceNum / sliceCount,
				((sliceNum+1).mod(sliceCount))/sliceCount
			]);

			// set effectLPFEnd to effectLPFStart if it is 0
			effectLPFEnd = Select.kr(effectLPFEnd < 0.001, [
				effectLPFEnd,
				effectLPFStart,
			]);
			// set effectTimestretch to beginning if it is 0
			effectTimestretchEnd = Select.kr(effectTimestretchEnd< 0.001,[
				effectTimestretchEnd,
				effectTimestretchStart,
			]);
			effectTimestretch = Line.kr(effectTimestretchStart, effectTimestretchEnd, effectTimestretchBeats*beatDuration);
			timestretchRate = Select.ar(effectTimestretch>1,[
				DC.ar(pos*frames),
				Phasor.ar(trig:TDelay.ar(Impulse.ar(0, 0),0),rate:rate/effectTimestretch,start:0,end:frames,resetPos:(pos*frames)).floor,
			]);


			// make sure retrigRateEnd is at least retrigRateStart
			retrigRateEnd = Select.kr(retrigRateEnd < retrigRateStart, [retrigRateEnd, retrigRateStart]);

			// if sliceDurationBeats = 0, make it infinite
			sliceDurationBeats = Select.kr(sliceDurationBeats < 0.001, [sliceDurationBeats, (seconds/sliceCount)/(60/bpmSource)]);

			// Calculate rate
			rate = rate*BufRateScale.ir(buf)*bpmTarget/bpmSource;
			// Calculate pitch
			rate = rate * (2 ** (pitch/12.0));

			// Calculate retriggers
			retrigCountFeedback = LocalIn.ar(1);
			retrigTrig = Impulse.ar(XLine.ar(retrigRateStart,retrigRateEnd,beatDuration*retrigRateChangeBeats) / beatDuration);
			retrigCount = PulseCount.ar(retrigTrig * (retrigCountFeedback<retrigNumTotal));
			sliceTrigger = sliceTrigger + (retrigTrig*(retrigCount<retrigNumTotal));

			// Increase pitch based on retrigger
			// rate = rate * (2 ** (retrigCount / 12));
			// Lower pitch based on retrigger
			rate = rate * (2 ** (retrigPitchChange * (retrigCount / 12)));
			rate = rate.abs;
			// Do reverse
			rate = Select.kr(effectReverse>0,[
				rate,
				rate * -1,
			]);

			// Change volume based on retrigger
			volumeDB = volumeDB + (A2K.kr(retrigCount) * retrigVolumeChange);


			// do timestretching
			sliceTrigger = sliceTrigger + ((effectTimestretch>1)*Impulse.ar(effectTimestretch/sliceSeconds));


			// Determine whether to toggle playback
			aOrB=ToggleFF.ar(sliceTrigger);
			crossfade=VarLag.ar(K2A.ar(aOrB),xfade,warp:\sine);

			posA=Phasor.ar(
				trig:(1-aOrB),
				rate:rate,
				end:frames,
				resetPos:timestretchRate,
			);
			posB=Phasor.ar(
				trig:aOrB,
				rate:rate,
				end:frames,
				resetPos:timestretchRate,
			);
			snd=(BufRd.ar(
				numChannels:2,
				bufnum:buf,
				phase:posA,
			)*crossfade)+(BufRd.ar(
				numChannels:2,
				bufnum:buf,
				phase:posB,
			)*(1-crossfade));

			snd = snd * Lag.kr(volumeDB.dbamp,0.2);

			// full envelope
			snd = snd * EnvGen.ar(Env.adsr(xfade,0.0,1.0,xfade,curve:\sine),gate,doneAction:2);

			// slice envelope
			snd = snd * EnvGen.ar(Env.new([0,1,1,0],[sliceAttackBeats,sliceDurationBeats,sliceReleaseBeats]*beatDuration,[-4,4]), sliceTrigger);


			// filtering 
			snd = RLPF.ar(snd, XLine.kr(effectLPFStart, effectLPFEnd, effectLPFBeats*beatDuration), 0.303);
			DetectSilence.ar(snd,time:1,doneAction:2);

			// Feedback
			LocalOut.ar(retrigCount);

			Out.ar(effectDryOut, snd*effectDry);
			Out.ar(effectCombOut, snd*XLine.kr(0.001,effectComb,sliceDurationBeats*beatDuration/2));
			Out.ar(effectReverbOut, snd*XLine.kr(0.001,effectReverb,sliceDurationBeats*beatDuration/2));
		}).add;

		SynthDef("out",{
			arg busReverb, busDry, busComb, volumeDB=0.0, reverbAmt=0.1, combAmt=0.0;
			var sndWet = In.ar(busReverb,2) * Lag.kr(volumeDB.dbamp,2);
			var sndDry = In.ar(busDry,2) * Lag.kr(volumeDB.dbamp,2);
			var sndComb = In.ar(busComb,2) * Lag.kr(volumeDB.dbamp,2);
			var snd = 				sndDry;

			// add in comb
			snd = snd + ((0.5*sndComb)+
					Pan2.ar(CombN.ar(sndComb[0], 0.02, Rand(0.01,0.06), Rand(1,2)),-1,0.9) +
					Pan2.ar(CombN.ar(sndComb[1], 0.02, Rand(0.01,0.06), Rand(1,2)),1,0.9));

			// add in reverb
			snd = SelectX.ar(
				reverbAmt,[
					snd,
					snd + Fverb.ar(sndWet[0],sndWet[1],200,tail_density: LFNoise2.ar(1/3).range(50,90),decay: LFNoise2.ar(1/3).range(50,70))
				]
			);

			snd = RHPF.ar(snd,60,0.303);
			snd = AnalogTape.ar(snd,
				bias: 0.5,
				saturation: 1.2,
				drive: 0.5,
				oversample: 2,
			);
			ReplaceOut.ar(0,snd*0.5);
		}).add;

		s.sync;
		~busDry = Bus.audio(s, 2);
		~busReverb = Bus.audio(s, 2);
		~busComb = Bus.audio(s, 2);
		s.sync;
		~synOut = Synth.tail(s,"out",[
			busReverb: ~busReverb,
			busDry: ~busDry,
			busComb: ~busComb,
			volumeDB: 6,
		]);
		s.sync;
		~sampleCache = Dictionary.new();
		~synthsPlaying = Dictionary.new();


		~playFromMsg = {
			arg msg, b;
			var synName = 1000000.rand.asString;
			var args = [
				\buf,b, 
				\effectDryOut, ~busDry,
				\effectCombOut, ~busComb,
				\effectReverbOut, ~busReverb,
			];
			// stop all currently playing synths
			~synthsPlaying.values.do({
				arg syn;
				if (syn.isPlaying,{
					if (syn.notNil,{
						syn.set(\gate,0);
					});
				});
			});
			// add msg[2:] array to args
			args = args ++ msg[2..];

			~synthsPlaying.put(synName,
				Synth.head(s,"sampler",args).onFree({
					[b,"freed"].postln;
					~synthsPlaying.removeAt(synName);
				});
			);

		};

		o = OSCFunc({ |msg|
			var filename = msg[1];
			msg.postln;
			if (~sampleCache.at(filename).isNil,{
				// load sample
				~sampleCache.put(filename, Buffer.read(s,filename,action:{ |b|
					["loaded",b].postln;
					~playFromMsg.(msg,b);
				}));
			},{
				~playFromMsg.(msg,~sampleCache.at(filename));
			});
		},'/sampler');
		["loaded",NetAddr.langPort, NetAddr.localAddr].postln;

		s.sync;
		Routine {
			inf.do({
				1.sleep;
				// print out the cpu usage
				["cpu usage", s.avgCPU, s.peakCPU].postln;
			});
		}.play;

		~bpmTarget = 180;
		~sliceNum = 32;
		Routine {
			inf.do({ arg beat;
				var sliceNum = beat.mod(~sliceNum);
				["playing",beat].postln;
				NetAddr.new("127.0.0.1", NetAddr.langPort).sendMsg("/sampler", "amen_5c2d11c8_beats16_bpm170.flac",
					\sliceCount, 32,
					\sliceNum, sliceNum,
					\sliceDurationBeats, 100,
					\bpmSource, 170,
					\bpmTarget, ~bpmTarget, 
					\effectTimestretchStart, 3.rand,
					\effectTimestretchEnd, 8.rand,
					\effectTimestretchBeats, 8.rand,
					\retrigNumTotal, 16.rand,
					\retrigRateStart, 16.rand,
					\retrigPitchChange,-1,
					\retrigVolumeChange,-2,
					\effectComb, (100.rand)/100.0,
				);

				(60/~bpmTarget/2).wait;
			});
		}.play;
	}.play;
});
)
