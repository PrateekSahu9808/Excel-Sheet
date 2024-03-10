import { RouterProvider, createBrowserRouter } from "react-router-dom";
import Home from "./Home";
import Sheet from "./Sheet";
import Load from "./Load";
import View from "./View";

const basename = "/Excel-Sheet"; // Replace with your desired base name

const router = createBrowserRouter(
  [
    {
      path: "/",
      element: <Home />,
    },
    {
      path: "/sheet",
      element: <Sheet />,
    },
    {
      path: "/load",
      element: <Load />,
    },
    {
      path: "/view/:key",
      element: <View />,
    },
  ],
  { basename }
);

const Route = () => {
  return <RouterProvider router={router} />;
};

export default Route;