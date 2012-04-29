/*PhDebug.js
*	Contains tiny functions for doing debugging output.
*/

/*function Lgf(text)
*	Writes a line to the logfile.
*/
function Lgf(text){ts.WriteLine(text);}

/*function Status(text)
*	Writes a line to the statusbox.
*/
function Status(text){
	var statbox = document.getElementById("statbox");
	statbox.value+= (text+"\n");
	statbox.scrollTop = statbox.scrollHeight;
}