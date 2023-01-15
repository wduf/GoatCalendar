<svelte:head>
	<link rel="icon" href="./favicon.png" type="image/png" />
	<title>GoatCalendar</title>
</svelte:head>

<main class="content-wrapper">
	<h1 class={"goatcalender-header"}><Icon icon="game-icons:ram-profile" style="color: rgba(200, 0, 0)"/>Calendar</h1>
	<div class={"file-drop-zone"} on:dragover|preventDefault on:drop|preventDefault={(e) => upload(e)}>
		<div class={"upload-icon-wrapper"}>
			{#if uploaded}
				<Icon icon="material-symbols:download-done-rounded" height="30%" />
				<p style="font-size: 18px; font-weight: bold">Schedule successfully uploaded.</p>
			{/if}
			{#if !uploaded}
				<Icon icon="material-symbols:upload-file-outline-rounded" height="30%" />
				<p style="font-size: 18px; font-weight: bold">Drag and drop schedule (.xlsx).</p>
			{/if}
		</div>
	</div>
	<div class={"buttons-wrapper"}>
		<button class="button-80" on:click={() => createICS()}>Download .ics File</button>
		<button class="button-80" on:click={() => openInstructions()}>Instructions</button>
	</div>
</main>

<style>
	*
	{
		margin: 0;
		padding: 0;
	}

	.content-wrapper
	{
		padding: 1px 0 0;
		min-height: 100%;
		background: url("/foisie.png") rgba(30, 30, 30, 0.6);
		background-size: cover;
		background-blend-mode: multiply;
		height: calc(100vh - 1px);
	}

	.goatcalender-header
	{
		margin: 1% 2% 0;
		display: flex;
		font-size: 50px;
		color: white;
		filter: drop-shadow(2px 2px 2px black)
	}

	.file-drop-zone
	{
		border: 3px solid black;
		margin: 1% 2% 2%;
		display: flex;
		justify-content: space-around;
		align-items: center;
		height: 70vh;
		border-radius: 20px;
		background-color: rgba(220, 220, 220, 0.85);
		backdrop-filter: blur(5px)
	}

	.upload-icon-wrapper
	{
		display: flex;
		flex-direction: column;
		align-items: center;
		gap: 15px;
	}

	.buttons-wrapper
	{
		display: flex;
		justify-content: space-between;
		margin: 0 2%;
	}

	/* CSS */
	.button-80
	{
		background-color: rgba(220, 220, 220, 0.85);
		backdrop-filter: blur(1px);
		backface-visibility: hidden;
		border-radius: .375rem;
		border-style: solid;
		border-width: .125rem;
		box-sizing: border-box;
		color: #212121;
		cursor: pointer;
		display: inline-block;
		font-family: Circular,Helvetica,sans-serif;
		font-size: 1.125rem;
		font-weight: 700;
		letter-spacing: -.01em;
		line-height: 1.3;
		padding: .875rem 1.125rem;
		position: relative;
		text-align: left;
		text-decoration: none;
		transform: translateZ(0) scale(1);
		transition: transform .2s;
		user-select: none;
		-webkit-user-select: none;
		touch-action: manipulation;
	}

	.button-80:not(:disabled):hover 
	{
		transform: scale(1.05);
	}

	.button-80:not(:disabled):hover:active
	{
		transform: scale(1.05) translateY(.125rem);
	}

	.button-80:focus
	{
		outline: 0 solid transparent;
	}

	.button-80:focus:before
	{
		content: "";
		left: calc(-1*.375rem);
		pointer-events: none;
		position: absolute;
		top: calc(-1*.375rem);
		transition: border-radius;
		user-select: none;
	}

	.button-80:focus:not(:focus-visible)
	{
		outline: 0 solid transparent;
	}

	.button-80:focus:not(:focus-visible):before
	{
		border-width: 0;
	}

	.button-80:not(:disabled):active
	{
		transform: translateY(.125rem);
	}
</style>

<script>
	import Icon from "@iconify/svelte";
	import readXlsxFile from "read-excel-file";
	import { createEvents } from "ics";

	var uploaded = null;

	function openInstructions()
	{
		window.open("https://github.com/wduf/goat-calendar/blob/main/README.md");
	}

	function upload(e)
	{
		if(e.dataTransfer.items)
		{
			let file = e.dataTransfer.items[0].getAsFile();
			if(file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
			{
				uploaded = file;
			}
			else
			{
				alert("invalid file type.")
				uploaded = null;
			}
		}
	}

	async function createICS()
	{
		if(!uploaded)
		{
			alert("please upload your schedule.");
			return;
		}
		let events = [];
		await readXlsxFile(uploaded).then((data) =>
		{
			for(let i = 3; i < data.length; i++)
			{
				let course = data[i];
				let listing = course[1];
				let delivery = course[6];
				let prof = course[9];
				let start = course[10];
				let end = parseDate(course[11]);
				let pattern = course[7];
				let patternArray = parseInfo(pattern);
				// meeting pattern exists
				// NOTE: if no meeting pattern (iqp, pqp, pc, etc.), don't add to calendar
				if(patternArray)
				{
					let days = parseDays(patternArray[1]);
					let times = parseTime(patternArray[2]);
					let start_dow = start.getDay();
					let off = getOffset(start_dow, days);
					// edge case: week starts after first meeting date
					start.setDate(start.getDate() + off);
					let startArray = parseDayAndTime(start, times[0]);
                    let endArray = parseDayAndTime(start, times[1]);
					let location = patternArray[3];
					let event = 
					{
						calName: "WPI Classes", 
						title: listing, 
						start: startArray,
						end: endArray,
						description: `${prof}, ${delivery}`,
						location: location,
						alarms: [ { action: "display", description: `15 minutes until ${listing}`, trigger: { minutes: 15, before: true } } ],
						recurrenceRule: generateRRule(days, end)
					};
					events.push(event);
				}
			}
		})
		const { error, value } = createEvents(events);
		if(error)
		{
			console.log(error);
			return;
		}
		// download .ics
		let cal = new File([value], "schedule.ics", { type: "text/calendar" });
		let dl_cal = document.createElement("a");
		dl_cal.download = "schedule";
		dl_cal.href = URL.createObjectURL(cal);
		dl_cal.click();
		// reset
		uploaded = null;
	}

	function getOffset(dow, days)
	{
		let day_vals = { "MO": 0, "TU": 1, "WE": 2, "TH": 3, "FR": 4 };
		let dows = [];  // days of week
		for(let i = 0; i < days.length; i++)
		{
			let val = day_vals[days[i]] - dow;
			if(val < 0)
			{
				val += 7
			}
			dows.push(val);
		}
		return Math.min(...dows);
	}

	function parseInfo(str)
	{
		if(!str)
		{
			return null;
		}
		let regex = /([MTWRF-]+) \| (\d{1,2}:\d{2} [AP]M - \d{1,2}:\d{2} [AP]M) \| (.*)/;
		return str.match(regex);
	}

	function parseDays(str)
	{
		let regex = /[MTWRF]/g;
		let arr = str.match(regex);

		const day = new Map();
		day.set("M", "MO");
		day.set("T", "TU");
		day.set("W", "WE");
		day.set("R", "TH");
		day.set("F", "FR");

		let ret = [];
		arr.forEach(element => ret.push(day.get(element)));
		return ret;
	}

	function parseTime(str){
		let regex = /(\d{1,2}:\d{2} [AP]M) - (\d{1,2}:\d{2} [AP]M)/;
		let arr = str.match(regex);

		let begin = convert12to24(arr[1]);
		let end = convert12to24(arr[2]);

		return [begin, end];
	}

	function convert12to24(time) {
        let i = time.indexOf(':');
        let j = time.indexOf(' ');

        let h = parseInt(time.substring(0, i));
        let m = time.substring(i, j);
        let s = time.substring(j + 1);

        if (s === "AM" && h === 12) {
            h = 0;
        } else if (s === "AM") {
            h = h;
        } else if (s === "PM" && h === 12) {
            h = 12;
        } else {
            h = h + 12;
        }
		return h.toString().concat(m);
    }

	function parseDate(day)
    {
        const pad = n => n < 10 ? `0${n}` : `${n}`;
        return [
            day.getFullYear(),
			pad(day.getMonth() + 1),
			pad(day.getDate() + 1),
			'T',
            pad(23),
            pad(59),
            pad(59),
            'Z'
        ].join('');
    }

    function parseDayAndTime(day, time)
    {
        let y = day.getFullYear();
        let m = day.getMonth() + 1;
        let d = day.getDate() + 1;

        let i = time.indexOf(':');
        let hr = parseInt(time.substring(0, i));
        let min = parseInt(time.substring(i+1));

        return [y, m, d, hr, min];
    }

	function generateRRule(days, end_date)
	{
		return `FREQ=WEEKLY;BYDAY=${days.join()};INTERVAL=1;UNTIL=${end_date}`;
	}
</script>
