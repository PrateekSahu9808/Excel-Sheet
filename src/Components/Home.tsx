import { Link } from "react-router-dom";
import "../index.css";
// import Image from "../aseets/Image/New-Excel-Blank-Workbook.png";
// import Img from "../aseets/Image/View.jpg";
const Home = () => {
  return (
    <div className="flex h-[full] bg-[#217246] gap-6">
      <div className="w-[30%] h-screen flex-cols  justify-center items-center gap-14 ">
        <div className="text-3xl font-bold mt-60 ml-7 text-white">
          Welcome to the Excel World!
        </div>
      </div>
      <div className="home w-[70%] flex flex-row items-center justify-center bg-[#F0F3FF]">
        <Link to="/Sheet">
          <img
            src="https://1.bp.blogspot.com/-hR_5BNiD8v4/XZlxc4_SFQI/AAAAAAAAG_E/mjH7uNgkZxgMo2kKc2IH80etGMG71MFeQCPcBGAYYCw/s1600/remove-excel-workbook-worksheet-password.png"
            alt="Create BlankWorkBook"
          />
        </Link>
        <Link to="/load">
          <img
            src="https://collab365.com/wp-content/uploads/2019/10/word-image-42.png"
            alt=""
          />
        </Link>
      </div>
    </div>
  );
};
export default Home;
