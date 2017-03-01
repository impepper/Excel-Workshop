<!DOCTYPE html>
<html>
<head>
    <title>Excel 投資決策輔助工具設計及製作</title>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="Draco is free PSD &amp; HTML template by @flamekaizar">
    <meta name="author" content="Afnizar Nur Ghifari">
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link rel="stylesheet" href="css/kube.min.css" />
    <link rel="stylesheet" href="css/font-awesome.min.css" />
    <link rel="stylesheet" href="css/custom.min.css" />
    <link rel="shortcut icon" href="img/favicon.png" />
	<link href='http://fonts.googleapis.com/css?family=Playfair+Display+SC:700' rel='stylesheet' type='text/css'>
	<link href='http://fonts.googleapis.com/css?family=Lato:400,700' rel='stylesheet' type='text/css'>
	<style>
		.intro h1:before {
			/* Edit this with your name or anything else */
			content: 'Excel';
		}
	</style>
</head>
<body>
	<!-- Navigation -->
	<div class="main-nav">
		<div class="container">
			<header class="group top-nav">
				<div class="navigation-toggle" data-tools="navigation-toggle" data-target="#navbar-1">
				    <span class="logo">Excel 投資決策輔助工具設計及製作</span>
				</div>
			    <nav id="navbar-1" class="navbar item-nav">
				    <ul>
				        <li class="active"><a href="#about">關於Excel</a></li>
				        <li><a href="#SyncChart">第一天</a></li>
				        <li><a href="#DailyData">第二天</a></li>
				        <li><a href="#DynammicData">第三天</a></li>
				    </ul>
				</nav>
			</header>
		</div>
	</div>

	<!-- Introduction -->
	<div class="intro section" id="about">
		<div class="container">
			<p>Hi,我是 Pepper,這裡是我所建置的Excel工作區</p>
			<div class="units-row units-split wrap">
				<div class="unit-20">
					<img src="img/Q_ava.png" alt="Avatar">
				</div>
				<div class="unit-80">
					<h1>透過 Excel，你可以：<br><span id="typed"></span></h1>					
				</div>
				<p>
					Excel是常見的試算表軟體，除了建立日常工作所需要的計算及統計表格之外，投資人也可以利用它豐富的函數庫以及極具彈性化的協同功能，進行關於投資方面的資料輸入、彙整、以及圖像話統計的工作，協助投資人在券商所提供的交易軟體之外，依循自己的投資理念，打造專屬於自己的投資工具～ 
				</p>
				<p>
					在這個臨時建立的網站之中，我將配合開設的課程，提供對於「<b>利用Excel打造自己的投資工具</b>」這個主題有興趣的夥伴們，取得課程當中部分的檔案資料，也希望各位夥伴們能夠在這三天的課程之中，了解並熟悉Excel在投資方面的應用
				</p>
			</div>
		</div>
	</div>

	<!-- 第一天，同步買賣超 -->
	<div class="work section second" id="SyncChart">
		<div class="container">
			<h1>第一天<br>外資/法人 同步買賣超資料的視覺表現 </h1>
			<!--<ul class="work-list">-->
			<ul class="award-list list-flat">
				<li>學習目標</li>
				<ul>
					<li>透過外部資料的連結功能，將所需要的當日買賣超標的網頁資訊置入到Excel工作表中</li>
					<li>利用COUNTIF函數進行資料的比對判斷</li>
					<li>利用格式化條件，快速透過顏色的改變，標示出不同類型的篩選標的</li>					
				</ul>

				<li>學習檔案</li>
				<ul>
					<li><a href="files/1-1.xlsm">基本檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
					<li><a href="files/1-1.xlsm">完成檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
				</ul>
		</div>
	</div>

	<!-- 第二天，盤後資料 -->
	<div class="work section second" id="DailyData">
		<div class="container">
			<h1>第二天<br>整理彙整每日的盤後資料</h1>
			<ul class="award-list list-flat">
				<li>學習目標</li>
				<ul>
					<li>透過外部資料的連結功能，將所需要的盤後數據資訊置入到Excel工作表中</li>
					<li>利用各項函數，計算出所需要之盤後數據變化</li>
					<li>製作圖表，以視覺化表現數據的變動狀況</li>
					<li>製作巨集，簡化工作流程</li>
				</ul>

				<li>學習檔案</li>
				<ul>
					<li><a href="files/1-1.xlsm">基本檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
					<li><a href="files/1-1.xlsm">完成檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
				</ul>
			</ul>
		</div>
	</div>

	<!-- 第三天，動態資料 -->
	<div class="work section second" id="DynamicData">
		<div class="container">
			<h1>第三天<br>當日 類股/個股 的即時強弱勢狀態</h1>
			<ul class="skill-list list-flat">
				<li>學習目標</li>
				<ul>
					<li>熟悉DDE以及RTD的觀念</li>
					<li>透過DDE以及RTD，動態連結券商軟體的資料到Excel之中，並且在Excel內進行即時的運算工作</li>
				</ul>

				<li>學習檔案</li>
				<ul>
					<li><a href="files/1-1.xlsm">基本檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
					<li><a href="files/1-1.xlsm">完成檔案</a>（基本工作表，已建立網路來源連結，不含公式）</li>
				</ul>
			</ul>
		</div>
	</div>

	<!-- Quote -->
	<div class="quote">
		<div class="container text-centered">
			<h1>希望讓自己的投資更加順利？讓Excel幫你一把</h1>
		</div>
	</div>

	<footer>
		<div class="container">
			<div class="units-row">
			    <div class="unit-50">
			    	<p>網站內容建置：Pepper</p>
			    </div>
			    <div class="unit-50">
					<ul class="social list-flat right">
						<li><a href="mailto:chihong.lin@gmail.com"><i class="fa fa-send"></i></a></li>
						<!--<li><a href="http://twitter.com/flamekaizar"><i class="fa fa-twitter"></i></a></li>-->
						<!--<li><a href="http://dribbble.com/flamekaizar"><i class="fa fa-dribbble"></i></a></li>-->
						<!--<li><a href="http://behance.com/flamekaizar"><i class="fa fa-behance"></i></a></li>-->
					</ul>
			    </div>
			</div>
		</div>
	</footer>

	<!-- Javascript -->
	<script src="js/jquery.min.js"></script>
	<script src="js/typed.min.js"></script>
    <script src="js/kube.min.js"></script>
    <script src="js/site.js"></script>
    <script>
		function newTyped(){}$(function(){$("#typed").typed({
		// Change to edit type effect
		strings: ["收集彙整籌碼資料", "提供交易決策輔助數據", "打造專屬即時監控環境"],

		typeSpeed:90,backDelay:700,contentType:"html",loop:!0,resetCallback:function(){newTyped()}}),$(".reset").click(function(){$("#typed").typed("reset")})});
    </script>
</body>
</html>