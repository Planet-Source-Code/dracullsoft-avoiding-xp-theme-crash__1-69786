Avoiding XP Theme Crash

Study of why vb6 applications with xp theme enabled sometimes crashes when closed. I have included an example of an application that crashed and two examples to show how to prevent it from crashing. Conclusions: It seems to be related to XP Theme not being closed correctly when a internal vb6 UserControl is on the form.
A simple way to fix it is to ensure it gets cleaned up either by placing an imagelist or a commond dlg control on the form.
Unzip the download and compile each of the 3 examples and study the source. I have tried to keep it as simple as possible.

/Dracull