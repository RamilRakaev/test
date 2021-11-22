class Hello extends React.Component {
    render() {
        return <h1>Привет, React</h1>;
    }
}
ReactDom.render(
    <Hello />,
    document.getElementById("content")
);