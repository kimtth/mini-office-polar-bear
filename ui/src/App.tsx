import { BrowserRouter, Route, Routes } from "react-router-dom";
import Home from "./apps/Home";
import OnlyOfficeEditor from "./apps/advanced/OnlyOfficeEditor";
import Writer from "./apps/writer/Writer";
import Slides from "./apps/slides/Slides";
import Calc from "./apps/calc/Calc";

export default function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<Home />} />
        <Route path="/advanced-editor" element={<OnlyOfficeEditor />} />
        <Route path="/writer" element={<Writer />} />
        <Route path="/slides" element={<Slides />} />
        <Route path="/calc" element={<Calc />} />
      </Routes>
    </BrowserRouter>
  );
}
