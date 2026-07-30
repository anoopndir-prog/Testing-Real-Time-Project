# Note to send with the .exe

Copy the message below into Teams/email alongside the file. Fill in the
two bracketed bits first.

---

Hi all,

Please help test the new Report Generator before we roll it out. It builds
the **Project Specification** and **Final Test Report** documents from the
usual source files, so you don't have to assemble them by hand.

**To run it:** download `SKF_Report_Generator.exe` and double-click. There
is nothing to install. It opens a normal app window - no browser, no
website, no setup. Your files never leave your machine.

**First launch takes a few seconds** with nothing on screen while it
unpacks itself. That's normal, not a hang - please don't double-click
repeatedly.

[IF THE BUILD IS UNSIGNED, KEEP THIS PARAGRAPH - OTHERWISE DELETE IT]
Windows will warn you about an "unknown publisher" because we don't have
our code-signing certificate yet. Click **More info** then **Run anyway**.
To confirm you have the genuine file, open Command Prompt and run:
`certutil -hashfile SKF_Report_Generator.exe SHA256`
It should print: [PASTE THE CHECKSUM FROM SKF_Report_Generator.exe.sha256.txt]

**You need Microsoft Word** for *Generate Report in PDF* - that's what
makes the PDF match our templates exactly. *Generate Report in Word*
works without it. Reports are saved to your **Downloads** folder.

**What I need you to check** - please try it on a job you've already done
by hand, so you can compare against a report you trust:

1. Does it open, and do both tools load?
2. Project Specification: attach the Excel request sheet, generate Word,
   and generate PDF.
3. Final Test Report: attach the Project Specification and the Test
   Inspection sheet, add monitoring sheets, generate Word and PDF.
4. Attach a seal photo and check the auto-crop picks the seal sensibly.
5. **Compare against your hand-made version.** Wrong or missing numbers
   matter far more than layout.

**Reporting back** - the useful details are:

- Which tool, and which button.
- What you attached (file types, roughly how many pages).
- What you expected vs what you got.
- The exact error text if one appeared - a screenshot is ideal.
- For wrong content: which section, and what it should have said.

Please **don't use this for a report going to a customer** until we've
finished testing. Keep doing those the current way.

Thanks,
[YOUR NAME]

---

## Before you send

- [ ] Built with `build_windows_exe.bat` (both security checks passed)
- [ ] Signed, or the unsigned paragraph above is kept and the checksum
      pasted in
- [ ] **You ran it yourself first**, end to end, on a real job - don't let
      the team find that it doesn't open
- [ ] Testers know which jobs to try it on
- [ ] `version_info.txt` version bumped, so you can tell which copy people
      are running when they report something

## Collecting the results

Ask everyone to reply in one thread rather than direct-messaging you.
Duplicate reports of the same fault are the fastest signal about what to
fix first, and you lose that if the feedback is scattered.
