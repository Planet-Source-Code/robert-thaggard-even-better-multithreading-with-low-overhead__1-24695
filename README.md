<div align="center">

## Even Better Multithreading with Low Overhead


</div>

### Description

Theres a code on PSC that says the best stable multithreading in vb6 is done with activex. I say thats the worst advice I've ever heard. This is one of my solutions.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Robert Thaggard](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/robert-thaggard.md)
**Level**          |Advanced
**User Rating**    |4.3 (103 globes from 24 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/robert-thaggard-even-better-multithreading-with-low-overhead__1-24695/archive/master.zip)





### Source Code

<p><b>Easy multithreading with low overhead - Part 1</b></p>
<p>Srideep Prasad posted an article on how to do safe multithreading in vb6 with
multi instancing. His "solution" required making an activex exe and
making new instances of it for each thread which obviously is very processor
consuming and defeats the very purpose of multithreading. His reason for this
"solution" was "hey, at least theres no more doevents." Give
me a break. I'm dont understand how he code made it to code of the month list.</p>
<p>My solution is simple and has low overhead.</p>
<p>1. Create an api dll using visual c++. If you dont know how to program c++,
thats no problem. You can use my template.<br>
2. Make a function that gets the address of the function you want to run in a
seperate thread.<br>
3. From here you can either use the dll as code running in the background to
serve as a "airbag" so you can call CreateThread safely from in the
dll, or you can call the function by yourself in the dll by address. (This is
called a callback routine. Many enumerated functions in the windows api do
this.)<br>
<br>
Part 1 of this tutorial will cover how to make a callback routine for your
multithreading.</p>
<p>The first step is to make a new Win32 Dynamic-Link Library workspace. Here is
my code template for an api dll.</p>
<pre>
</font><font face="Courier" size="2"><font color="#0000FF">#include</font> &lt;windows.h&gt;
<font color="#008000">// This may be a little confusing to some people.
// All this next line does is specify a vb-safe calling convention
// CALLBACK* says that the variable type is actually a function, in this case a vb function
// THREADED_FUNC is the variable type that the function will be called in the dll. I could have put anything else in here
// typedef BOOL means that the function has a return value of boolean
// (int) means that the function has one paramater and its an integer. You could put as many of these as you need, depending
// 	on the number of parameters your function takes. ie your function takes an integer and two strings. You would put
//	(int, LPCSTR, LPCSTR)</font>
<font color="#0000FF">typedef</font> BOOL (CALLBACK* THREADED_FUNC) (<font color="#0000FF">int</font>);
<font color="#008000">// Function prototypes</font>
<font color="#0000FF">void</font> FreeProcessor(<font color="#0000FF">void</font>);
LONG <font color="#0000FF">__declspec</font>(<font color="#0000FF">dllexport</font>) WINAPI MakeThread(THREADED_FUNC &amp;pTFunc, <font color="#0000FF">int</font> nPassedValue);
<font color="#008000">// Starting and ending point of the dll, required for every api dll</font>
<font color="#0000FF">extern</font> &quot;C&quot; <font color="#0000FF">int</font> APIENTRY DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	<font color="#0000FF">if</font> (dwReason == DLL_PROCESS_ATTACH)
	{
		<font color="#008000">// dll starts processing here
		// inital values and processing should go here</font>
	}
	<font color="#0000FF">else if</font> (dwReason == DLL_PROCESS_DETACH)
	{
		<font color="#008000">// dll stops processing here
		// all clean up code should go here</font>
	}
<font color="#0000FF">	return</font> 1;
}
<font color="#008000">// MakeThread - Function that calls function by address (This is the callback routine)
// This function accepts a THREADED_FUNC which is actually the address of the threaded function
// It also accepts the parameters your function takes which is an integer for this example. You will need to set the
//	number of parameters to match the function you wrote</font>
LONG <font color="#0000FF">__declspec</font>(<font color="#0000FF">dllexport</font>) WINAPI MakeThread(<font color="#0000FF">int</font> nPassedValue, THREADED_FUNC &amp;pTFunc)
{
	<font color="#008000">// try-catch block for error handling</font>
	<font color="#0000FF">try</font>
	{
		<font color="#0000FF">do</font>
		{
<font color="#008000">			// call the function by address and examin return value
</font>			<font color="#0000FF">if</font> (pTFunc(nPassedValue) == FALSE)
				<font color="#0000FF">return</font> 1;
			FreeProcessor();
		} <font color="#0000FF">while</font> (<font color="#0000FF">true</font>);
	}
	<font color="#0000FF">catch</font> (<font color="#0000FF">int</font>) { <font color="#0000FF">return</font> 0; }
}
<font color="#008000">// FreeProcessor function written by Jared Bruni
</font><font color="#0000FF">void</font> FreeProcessor(<font color="#0000FF">void</font>)
{
	MSG Msg;
	<font color="#0000FF">while</font>(PeekMessage(&amp;Msg,NULL,0,0,PM_REMOVE))
	{
		<font color="#0000FF">if</font> (Msg.message == WM_QUIT)<font color="#0000FF">break</font>;
		TranslateMessage(&amp;Msg);
		DispatchMessage(&amp;Msg);
	}
}
</font></pre>
<br>
The next step is to create a export definitions file for MakeThread. This
is very simple.
<p><font face="Courier" size="1"><font color="#FF0000">LIBRARY MyFile</font><br>
DESCRIPTION 'Callback multithreading dll for MyProgram'<br>
CODE PRELOAD MOVEABLE DISCARDABLE<br>
DATA PRELOAD MOVEABLE SINGLE<br>
<br>
HEAPSIZE 4096<br>
EXPORTS<br>
    MakeThread @1<br>
</font></p>
<p>I highlighted the LIBRARY line for a good reason. Make sure whatever you type
after LIBRARY is the name of the cpp file that your DllMain is in. For example
if your DllMain is in a file called "BigLousyDll.cpp", then you would
type LIBRARY BigLousyDll<br>
<br>
Also make sure that the export definitions file is the same name as the cpp file
your DllMain is in. Like I said, if your DllMain is in a file called "BigLousyDll.cpp",
you would name your export definitions file BigLousyDll.def<br>
<br>
Once you compile your dll, it should automatically be registered. I would put it
in your system or system32 folder so you don't have to type a explicit path to
it in your vb file.</p>
<p><font face="Courier" size="1"><font color="#0000FF">Public Declare Function</font>
MakeThread <font color="#0000FF">Lib</font> "MyFile.dll" (lpCallback <font color="#0000FF">As
Any</font>, <font color="#0000FF">ByVal</font> nInt <font color="#0000FF">As
Integer</font>) <font color="#0000FF">As Long<br>
Public </font>i<font color="#0000FF"> As Integer<br>
Public </font>nTimes<font color="#0000FF"> As Integer<br>
<br>
Public Function </font>MyFunction(ByVal nValue As Integer) As Boolean<br>
    nTimes = nTimes + 1<br>
<font color="#0000FF">   </font> <font color="#0000FF">If</font>
nTimes > 0 <font color="#0000FF">Then</font><br>
        <font color="#0000FF">If</font> i
< 20 <font color="#0000FF">Then</font><br>
            i = i + 1<br>
        <font color="#0000FF">End If</font><br>
        MyFunction = <font color="#0000FF">True   
</font><font color="#008000">'Tells dll to keep running through function</font><br>
        <font color="#0000FF">Exit Function<br>
    Else</font><br>
<font color="#0000FF">   </font>     i = nValue<br>
        MyFunction = <font color="#0000FF">True   
</font><font color="#008000">'Tells dll to keep running through function</font><font color="#0000FF"><br>
        Exit Function<br>
   </font> End If<font color="#0000FF"><br>
    </font>MyFunction =<font color="#0000FF">
False    </font><font color="#008000">'Tells dll to stop</font><font color="#0000FF"><br>
End Function<br>
<br>
Sub </font>Main()<br>
    <font color="#0000FF">If Not</font> MakeThread(<font color="#0000FF">AddressOf</font>
MyFunction, 3) <font color="#0000FF">Then</font><br>
        MsgBox "Multithreading
error"<br>
    <font color="#0000FF">Else</font><br>
        MsgBox "Success"<br>
    <font color="#0000FF">End If</font><br>
<font color="#0000FF">End Sub</font></font></p>
<p><br>
If you find this code helpful, vote only if you want to. I dont care if I win
coding contest. I just thought this solution is excellent compared to what
Srideep Prasad posted.</p>

