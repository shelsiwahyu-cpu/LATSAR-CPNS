import { useState, useEffect } from 'react';
import Select from 'react-select'; // pastikan sudah npm install react-select

function App() {
    const [datas, setDatas] = useState([])
    const [userSelect, setUserSelect] = useState("")
    const [isShow, setIsShow] = useState(false)

    const getBerries = async () => {
        try {
            // Ganti dengan API yang benar, contoh: PokeAPI untuk berries
            const berries = await fetch("https://pokeapi.co/api/v2/berry")
            const value = await berries.json()
            let result = value.results.map(data => {
                return {
                    label: data.name,
                    value: data.name
                }
            })
            setDatas(result) // SET STATE - ini yang kurang!
        } catch (error) {
            console.error("Error fetching data:", error)
        }
    }

    useEffect(() => { 
        getBerries()
    }, [])

    const handleSubmit = () => {
        setIsShow(state => !state)
    }

    const handleChange = (value) => {
        setUserSelect(value)
    }

    return (
        <div className="App"> {/* Perbaikan typo */}
            <h1>{isShow ? userSelect : ""}</h1>
            <button onClick={() => handleSubmit()}>
                {isShow ? "Hide Button" : "Search"}
            </button>
            <br />
            <br />
            <Select 
                options={datas} 
                onChange={(e) => handleChange(e.value)}
            />
        </div>
    );
}

export default App;