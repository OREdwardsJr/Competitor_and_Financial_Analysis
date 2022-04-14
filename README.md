# Competitor and Financial Analysis
This project was created to assist a company in gathering a more thorough and accurate analysis of its cost and returns. While the code may be tweaked to specific use-cases, it is not intended to be used outside of its original intent. This repository is mostly being utilized as an area to display the project. However, contributions and/or recommendations are welcomed.


<div id="top"></div>

[![LinkedIn][linkedin-shield]][linkedin-url]



<!-- PROJECT LOGO 
<br />
<div align="center">
  <a href="https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis">
    <img src="images/logo.png" alt="Logo" width="80" height="80">
  </a> -->

<h3 align="center">Competitor and Financial Analysis</h3>

  <p align="center">
    Using Python and a variety of libraries (primarily: Openpyxl, Pandas, Requests), this project analyzes a wholesaling company's cost-of-goods by scraping and comparing it to public information about its products.
    <br />
    <a href="https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis">View Demo</a>
    ·
    <a href="https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/issues">Report Bug</a>
    ·
    <a href="https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/issues">Request Feature</a>
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>
    <li><a href="#acknowledgments">Acknowledgments</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

<!-- [![Product Name Screen Shot][product-screenshot]](https://example.com) -->

This project was created because I realized that the company could receive significant improvement in their financial analysis through the usage of Python. The company was utilizing Microsoft Excel to manage its databases. This unfortunately limited their ability to effectively track and manage their data.

A 12-month restrospective investigation into various pricing and purchasing data is conducted. It analyzes data such as: purchasing cost, selling price, estimated retail pricing, and etc. This project was able to uncover that the company was missing target retail estimations by 35-65% on average, in respect to retail and sale prices.

Through this project, we were able to reconstruct an algorithm that was able to be within 3-7% of estimated retail and sale prices.

<p align="right">(<a href="#top">back to top</a>)</p>



### Built With

* [Python](https://www.python.org/)
* - [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
* - [Pandas](https://pandas.pydata.org/)
* - [Requests](https://docs.python-requests.org/en/latest/)

<p align="right">(<a href="#top">back to top</a>)</p>




<!-- USAGE EXAMPLES -->
## Usage

This project has a structure that can be replicated if you are dealing with Excel databases and performing HTTP requests that do not require Javascript execution. In its current state, it will not execute outside of the expected use cases without tweaking.

The code has had sensitive data, like company names and file paths, redacted. 

If you'd like to use this project for whatever reason then please contact me to discuss tailoring the code to your specific needs.

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- ROADMAP -->
## Project Structure

* Section 0
*  - Imports / Initial variables
* Section 1
*  - Parses the child company's database to analyze against
* Section 2
*  - Scrapes parent company to obtain selling prices
* Section 3
*  - Sanitizes data
* Section 4
*  - Performs calculations
* Section 5
*  - Sanitizes data
* Section 6
*  - Save to Excel
* Debug
*  - Demonstrates issues that I ran into and how I addressed them.

See the [open issues](https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".
Don't forget to give the project a star! Thanks again!

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#top">back to top</a>)</p>



<!-- CONTACT -->
## Contact

Orlando Edwards Jr. - [LinkedIn](https://linkedin.com/in/orlando-edwards-jr) - oredwardsjr@gmail.com

Project Link: [https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis](https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis)

<p align="right">(<a href="#top">back to top</a>)</p>





<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/OREdwardsJr/Competitor_and_Financial_Analysis.svg?style=for-the-badge
[contributors-url]: https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/OREdwardsJr/Competitor_and_Financial_Analysis.svg?style=for-the-badge
[forks-url]: https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/network/members
[stars-shield]: https://img.shields.io/github/stars/OREdwardsJr/Competitor_and_Financial_Analysis.svg?style=for-the-badge
[stars-url]: https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/stargazers
[issues-shield]: https://img.shields.io/github/issues/OREdwardsJr/Competitor_and_Financial_Analysis.svg?style=for-the-badge
[issues-url]: https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/issues
[license-shield]: https://img.shields.io/github/license/OREdwardsJr/Competitor_and_Financial_Analysis.svg?style=for-the-badge
[license-url]: https://github.com/OREdwardsJr/Competitor_and_Financial_Analysis/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/orlando-edwards-jr/
[product-screenshot]: images/screenshot.png
