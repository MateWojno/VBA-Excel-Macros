<div id="about">
        <h1 align="center">VBA Excel Macros for my Company</h1>
        <p align="center">author @MateWojno, mateusz.k.wojno@gmail.com <br>Start   17-10-2022<br>End 19-10-2022</p>
</div>
<div id="toc"> 
        <h1 align="center">Table of content</h1>
        <ul align="center">
                <li><a href="#about">About</a></li>
                <li><a href="#res">Resources</a></li>
                <li><a href="#extensions">Extensions</a></li>
                <li><a href="#algorithms">Algorithms</a></li>
                <li><a href="#api">API</a></li>
        </ul>
</div>
<div id="res"> 
        <h1 align="center">Resources:</h1>
        <ul>
                <li><a href="https://www.wallstreetmojo.com/vba-rename-sheet/">VBA coding</a></li>
                <li><a href="https://www.wallstreetmojo.com/macros-in-excel/">Macros in Excel</a></li>
                <li><a href="https://file.org/extension/bas#:~:text=BASIC%20is%20a%20programming%20language%20that%20was%20created,language%2C%20it%20is%20saved%20with%20the.bas%20file%20extension.">.bas file extension</a></li>
                <li><a href="https://www.wallstreetprep.com/self-study-programs/the-ultimate-excel-vba-course/">Paid VBA course</a></li>
                <li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/resources/power-query-differences">About Power Query</a></li>
                <li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/resources/vba-differences">Differences between VBA Macros and Office Scripts(online)</a></li>
                <li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/">Office Scripts documentation</a></li>
                <li><a href="https://en.wikipedia.org/wiki/Microsoft_Access">About MS Access</a></li>
                <li><a href="https://www.lifewire.com/mdb-file-2621974">What Is an MDB File?</a></li>
                <li><a href="https://learn.microsoft.com/en-us/office/dev/scripts/develop/script-buttons?source=recommendations">About scripted buttons in Microsoft Excel Desktop App</a></li>
                <li><a href="https://www.exceldome.com/solutions/rename-an-active-excel-worksheet/#:~:text=VBA%20Methods%3A%20Using%20VBA%20you%20can%20rename%20an,worksheet%20and%20you%20can%20then%20rename%20the%20worksheet">How to rename worksheet</a></li>
        </ul>
</div>
<div id="extensions">
        <h1 align="center">Tools/extensions required</h1>
        <ul> </ul>
                <li>Microsoft Excel 2019 (pro recommended)</li>
                <li>Power Query (built in)</li>
                <li>Visual Studio Code (recommended for Devs)</li>
                <li>XVBA - Live Server VBA</li>
                <li>VBA v0.6.0 serkonda7</li>
                <li>vba-snippets Scott Spence</li>
                <li>Dedicated App built by me (optional)</li>
                <li>for old .mdb files you need MS Access 2010</li>
        <a href="#toc">Return to the top</a>
</div>
<div id="algorithms">
        <div id="reading-database">
                <h1 align="center">#1/ Reading database file algorithm</h1>
                <h2>Input data:</h2>
                <ul>
                        <li>MS Access database file</li>
                        <li>.mdb/.accdb</li>
                </ul>       
                <h2>Description</h2>
                <ol> 
                        <li>Start</li>
                        <li>Fetch data from given path</li>
                        <li>Add new worksheet, set it active</li>
                        <li>Rename active worksheet as "wdb"</li>
                        <li>(Optional - send logs)</li>
                        <li>Stop</li>
                </ol>
                <h2>Output data</h2>
                <ol> 
                        <li>Excel worksheet</li>
                        <li>.xlsm - macros enabled</li>
                </ol>
                <a href="#toc">Return to the top</a>
                </div>
        <div id="transform-database">
                <h1 align="center">#2/ Data transformation in power query algorithm</h2>
                <h2>Input data</h2>
                        <p>Output from previous algorithm, active excel sheet</p>
                <h2>Description</h2>
                <ol>
                        <li>Start</li>
                        <li>Fetch previous data</li>
                        <li>Transform data from worksheet "wdb" in given range</li>
                        <li>(Optional - send logs)</li>
                        <li>Stop</li> 
                </ol>        
                <h2>Output data</h2>
                        <p>Properly transformed table</p>
                <a href="#toc">Return to the top</a>
        </div>
        <div id="refresh">
                <h1 align="center">#3/ Refresh loop algorithm</h1>
                <h2>Input data</h2>
                        <p>Active worksheet named "wdb"</p>
                <h2>Description</h2>
                <ol>
                        <li>Start</li>
                        <li>Delete specified query connection </li>
                        <li>Delete specified "wdb" worksheet</li>
                        <li>(Optional - send logs)</li>
                        <li>Stop</li> 
                </ol> 
                <h2>Output data</h2>
                        <p>None</p>
                <a href="#toc">Return to the top</a>
        </div>
        </div>
        <div id="api" align="center">
        <h1 align="center">#3/ API inside Excel</h1>
       <img src="https://user-images.githubusercontent.com/110040191/197168722-4ef0c86d-bd2d-4d7d-b130-aab0cf673538.png" alt="API">
        <h2>First button does the job of <a href="#reading-database">algorithm #1</a> and <a href="#transform-database">algorithm #2</a><br>
        Second button refreshes and deletes the query and "wdb" sheet - <a href="#transform-database">algorithm #3</a><br>
        Last button displays <a href="#res">info.</a><hr>
        <h5>Unfortunately, I can't share more information about this project due to the company's NDA.</h5>
        <a href="#toc">Return to the top</a>
        </div>





