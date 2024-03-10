import { RouterProvider, createBrowserRouter } from "react-router-dom";
import Home from "./Home";
import Sheet from "./Sheet";
import Load from "./Load";
import View from "./View"; // Import the View component

const router = createBrowserRouter([
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
]);
const Route = () => {
  return <RouterProvider router={router} />;
};
export default Route;
