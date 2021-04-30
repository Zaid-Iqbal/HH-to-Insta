<!DOCTYPE html>
<html>

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>README</title>
  <link rel="stylesheet" href="https://stackedit.io/style.css" />
</head>

<body class="stackedit">
  <div class="stackedit__html"><h1 id="hardware-hub-to-instagram">Hardware Hub to Instagram</h1>
<p>An Instagram bot that allows me to automatically post about a product from the Hardware Hub website (<a href="http://zed.exioite.com">zed.exioite.com</a>).</p>
<h2 id="goal">Goal</h2>
<p>Increase traffic to the Hardware Hub website</p>
<h2 id="technologies">Technologies</h2>
<p>-InstagramApiSharp: Instagram API for C#</p>
<blockquote>
<p><a href="https://github.com/ramtinak/InstagramApiSharp/">https://github.com/ramtinak/InstagramApiSharp/</a></p>
</blockquote>
<p>-NPOI: Allows me to read and write excel files in C#.</p>
<blockquote>
<p><a href="https://github.com/nissl-lab/npoi">https://github.com/nissl-lab/npoi</a></p>
</blockquote>
<h2 id="procedure">Procedure</h2>
<ol>
<li>getNextID();</li>
</ol>
<ul>
<li>Go through Twitter.xlsx and find the last product that was tweeted. Grab that product’s ID and return it</li>
</ul>
<ol start="2">
<li>Post(ID);</li>
</ol>
<ul>
<li>Use the ID to find the body of the text so be added to the Instagram post and to find the photo in hard drive</li>
</ul>
<h2 id="improvements">Improvements</h2>
<ol>
<li>Find a way to automatically run this program at 4:00 pm CST every day (the time with most US instagram traffic) without needing to have my laptop on at that time</li>
</ol>
<blockquote>
<p>sometime I forget to keep my laptop on at 4 pm and it doesn’t post and I completely forget</p>
</blockquote>
<ol start="2">
<li>Set up an SQL database on my laptop so the data is pulled from and stored in a database instead of an excel file</li>
</ol>
</div>
</body>

</html>
