import React, { useState, useEffect } from "react";
import {
    Row,
    Col,
    Badge,
    Button,
    Card,
    CardHeader,
    CardBody,
    Table,
    Container,
    Modal,
    ModalHeader,
    ModalBody,
    ModalFooter,
    Form,
    FormGroup,
    Label,
    Input,
    Spinner,
    UncontrolledDropdown,
    DropdownToggle,
    DropdownMenu,
    DropdownItem,
    Alert,
} from "reactstrap";
import Admin from "layouts/Admin";
import { useSession } from "next-auth/react";
import FormData from "form-data";
import { useRouter } from "next/router";
import path from "path";



function CSV() {
    const router = useRouter();

    const [modalOpen, setModalOpen] = useState(false);
    const [screenModal, setScreenModal] = useState(false);
    const [trackModal, setTrackModal] = useState(false);
    const [positions, setPositions] = useState([]);
    const [csvFiles, setCsvFiles] = useState([]);
    const [vessels, setVessels] = useState([]);
    const { data: session, status } = useSession();
    const [fileFormData, setFileFormData] = useState(new FormData());
    const [loading, setLoading] = useState(false);
    const [errorAlert, setErrorAlert] = useState(false);
    const [file, setFile] = useState(null);
    const [reportType, setReportType] = useState('');
    const [reportingPeriod, setReportingPeriod] = useState('');
    const [businessUnit, setBusinessUnit] = useState('');
    const [csvViewModalOpen, setCsvViewModalOpen] = useState(false);
    const [csvData, setCsvData] = useState(null);
    const [selectedVessel, setSelectedVessel] = useState(null);

    useEffect(() => {
        if (session) {
            fetchCsvFiles();
        }
    }, [session]);

    const [currentView, setCurrentView] = useState('Bordereaux Reporting');

    const renderTableContent = () => {
        switch (currentView) {
            case 'Bordereaux Reporting':
                return renderBordereauxTable();
            case 'Vessel Database':
                return renderVesselTable();
            case 'Signing':
                return renderSigningTable();
            default:
                return null;
        }
    };

    const renderBordereauxTable = () => (
        <>
            <thead className="thead-light">
                <tr>
                    <th scope="col">Reporting Period</th>
                    <th scope="col">Business Unit</th>
                    <th scope="col">Upload Date</th>
                    <th scope="col">Status</th>
                    <th scope="col" className="text-end">
                        <Button size="sm" color="primary" onClick={() => setModalOpen(true)}>
                            Upload CSV
                        </Button>
                    </th>
                </tr>
            </thead>
            <tbody>
                {Array.isArray(csvFiles) && csvFiles.map((file, index) => (
                    <tr key={index}>
                        <td>{file.reportingPeriod}</td>
                        <td>{file.businessUnit}</td>
                        <td>{new Date(file.uploadedAt).toLocaleString()}</td>
                        <td>
                            <Badge color={file.status === 'processed' ? 'success' : 'warning'}>
                                {file.status}
                            </Badge>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">

                            <UncontrolledDropdown>
                                <DropdownToggle
                                    className="btn-icon-only text-light"
                                    role="button"
                                    size="sm"
                                    color=""
                                    onClick={(e) => e.preventDefault()}
                                >
                                    <i className="fas fa-ellipsis-v" />
                                </DropdownToggle>
                                <DropdownMenu className="dropdown-menu-arrow" right>
                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "EUR_statement_of_accounts.csv") && file._files.some(f => f.fileName === "premiumsummary.csv") && (
                                        <DropdownItem onClick={() => {

                                            getSummary(file._id, "EUR_statement_of_accounts.csv");
                                        }}>
                                            Generate EUR Statement
                                        </DropdownItem>
                                    )}
                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "statement_of_accounts.csv") && file._files.some(f => f.fileName === "premiumsummary.csv") && (
                                        <DropdownItem onClick={() => {

                                            getSummary(file._id, "statement_of_accounts.csv");
                                        }}>
                                            Generate USD Statement
                                        </DropdownItem>
                                    )}
                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "claimsummary.csv") && (
                                        <>
                                            <DropdownItem onClick={() => {
                                                getSummary(file._id, "claimsummary.csv");
                                            }}>
                                                View Claim Summary
                                            </DropdownItem>
                                        </>
                                    )}

                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "premiumsummary.csv") && (
                                        <>
                                            <DropdownItem onClick={() => {
                                                getSummary(file._id, "premiumsummary.csv");
                                            }}>
                                                View Premium Summary
                                            </DropdownItem>
                                        </>
                                    )}


                                </DropdownMenu>


                            </UncontrolledDropdown>
                        </td>
                    </tr>
                ))}
            </tbody>
        </>
    );

    const renderVesselTable = () => (
        <>
            <thead className="thead-light">
                <tr>
                    <th scope="col">Vessel Name</th>
                    <th scope="col">IMO Number</th>
                    <th scope="col">Status</th>
                    <th scope="col" className="text-end">
                    </th>
                </tr>
            </thead>
            <tbody>
                {vessels.map((vessel, index) => (
                    <tr key={index}>
                        <td>{vessel.vesselName}</td>
                        <td>{vessel.imoNumber}</td>
                        <td>
                            <Badge color={(() => {
                                if (!vessel.OFAC) return 'default';
                                if (vessel.OFAC?.results?.[0]?.matchCount > 0) return 'warning';
                                return 'success';
                            })()}>
                                {(() => {
                                    if (!vessel.OFAC) return 'Pending Screening';
                                    if (vessel.OFAC?.results?.[0]?.matchCount > 0) return 'Matches Found';
                                    return 'Active';
                                })()}
                            </Badge>
                        </td>
                        <td className="text-end">
                            <UncontrolledDropdown>
                                <DropdownToggle
                                    className="btn-icon-only text-light"
                                    role="button"
                                    size="sm"
                                    color=""
                                    onClick={(e) => e.preventDefault()}
                                >
                                    <i className="fas fa-ellipsis-v" />
                                </DropdownToggle>
                                <DropdownMenu className="dropdown-menu-arrow" right>
                                    <DropdownItem
                                        onClick={() => {
                                            setSelectedVessel(vessel);
                                            setScreenModal(true);
                                        }}
                                    >
                                        Screening
                                    </DropdownItem>
                                    <DropdownItem
                                        onClick={() => {
                                            trackVessel(vessel)

                                        }}
                                    >
                                        View Current Location
                                    </DropdownItem>
                                    <DropdownItem>
                                        View Linked Premiums
                                    </DropdownItem>
                                    <DropdownItem>
                                        View Linked Claims
                                    </DropdownItem>
                                </DropdownMenu>
                            </UncontrolledDropdown>
                        </td>
                    </tr>
                ))}
            </tbody>
        </>
    );

    const renderSigningTable = () => (
        <>
            <thead className="thead-light">
                <tr>
                    <th scope="col">Document Name</th>
                    <th scope="col">Uploaded By</th>
                    <th scope="col">Upload Date</th>
                    <th scope="col">Signing Status</th>
                    <th scope="col" className="text-end">
                        <Button size="sm" color="primary" onClick={() => setModalOpen(true)}>
                            Upload Document
                        </Button>
                    </th>
                </tr>
            </thead>
            <tbody>
                {/* Signing table body content */}
            </tbody>
        </>
    );


    const fetchCsvFiles = async () => {
        try {
            const [csvResponse, vesselsResponse] = await Promise.all([
                fetch(`${process.env.API_BASE_URL}/api/getCsvFiles`, {
                    headers: {
                        Authorization: `Bearer ${session.accessToken}`,
                    },
                }),
                fetch(`${process.env.API_BASE_URL}/api/getVessels`, {
                    headers: {
                        Authorization: `Bearer ${session.accessToken}`,
                    },
                })
            ]);

            const files = await csvResponse.json();
            const vessels = await vesselsResponse.json();

            setCsvFiles(files);
            setVessels(vessels);
            console.log("Vessel", vessels);
        } catch (error) {
            console.error("Error fetching CSV files:", error);
        }
    };

    const handleFileChange = async (event) => {
        const file = event.target.files[0];
        setFile(file);

    };

    const handleFileUpload = async (event) => {
        event.preventDefault();
        const formData = new FormData();
        setLoading(true);
        if (file) {
            formData.append('file', file);
            console.log("File appended:", {
                name: file.name,
                type: file.type,
                size: file.size,
                lastModified: new Date(file.lastModified).toISOString()
            });
        } else {
            console.log("No file selected");
        }

        formData.append('reportType', reportType);
        formData.append('reportingPeriod', reportingPeriod);
        formData.append('businessUnit', businessUnit);
        console.log("FormData contents:");
        for (let [key, value] of formData.entries()) {
            if (key === 'file') {
                console.log(`${key}: [File object]`);
            } else {
                console.log(`${key}: ${value}`);
            }
        }

        let apiEndpoint;
        apiEndpoint = `${process.env.API_BASE_URL}/api/generateSummary`;


        try {
            const response = await fetch(apiEndpoint, {
                method: 'POST',
                headers: {
                    Authorization: `Bearer ${session.accessToken}`,
                },
                body: formData,
            });
            if (response.ok) {
                setModalOpen(false);
                fetchCsvFiles();
            } else {
                console.log(response.message, "::MSG")
                setErrorAlert(true);
                //    setModalOpen(!modalOpen)
            }
        } catch (error) {
            console.error("Error uploading file:", error);
        } finally {
            setLoading(false);
        }
    };

    const getSummary = async (id, fileName) => {
        setLoading(true);
        let nFileName = encodeURIComponent(fileName);
        try {
            const res = await fetch(
                `${process.env.API_BASE_URL}/api/getattachments?id=${encodeURIComponent(id)}&name=${nFileName}`,
                {
                    headers: {
                        Authorization: `Bearer ${session.accessToken}`,
                    },
                }
            );
            const text = await res.text();
            const parsedData = parseCSV(text);
            setCsvData(parsedData);
            setCsvViewModalOpen(true);
        } catch (error) {
            console.error("Error fetching CSV:", error);
        } finally {
            setLoading(false);
        }
    };

    const trackVessel = async (vessel) => {
        // event.preventDefault();
        setLoading(true);
        fetch(`${process.env.API_BASE_URL}/api/trackVessel`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                Authorization: `Bearer ${session.accessToken}`,
            },
            body: JSON.stringify(vessel),
        })
            .then((res) => res.json())
            .then((res) => {
                if (res.data && res.data.positions) {
                    setPositions(res.data.positions);
                    setTrackModal(true);
                    setDefaultAlert(false);
                } else {
                    // Handle the error or set a default value
                    setPositions([]); // Setting an empty array as default value
                }
                console.log(res, "::RESS");
            })
            .catch((error) => console.error(error))
            .finally(() => {
                setLoading(false);

            });
    };

    const screenVessel = async (vessel) => {
        console.log("Screening vessel:", vessel);
        setLoading(true);
        try {
            const res = await fetch(
                `${process.env.API_BASE_URL}/api/screenVessels`,
                {
                    headers: {
                        Authorization: `Bearer ${session.accessToken}`,
                    },
                    method: 'POST',
                    body: JSON.stringify(vessel),
                }
            );

        } catch (error) {
            console.error("Error fetching CSV:", error);
        } finally {
            fetchCsvFiles();
            setLoading(false);
            setScreenModal(false);
        }
    };

    const parseCSV = (text) => {
        const lines = text.split('\n');
        const headers = lines[0].split(',');
        const data = lines.slice(1).map(line => {
            const values = line.split(',');
            return headers.reduce((obj, header, index) => {
                obj[header.trim()] = values[index];
                return obj;
            }, {});
        });
        return { headers, data };
    };

    if (status === "loading") {
        return <p>Loading...</p>;
    }

    if (status === "unauthenticated") {
        router.push("/auth/login");
        return null;
    }

    return (
        <Admin>
            {loading && (
                <div
                    style={{
                        position: "absolute",
                        top: 0,
                        left: 0,
                        width: "100vw",
                        height: "100vh",
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        backgroundColor: "rgba(0,0,0,0.5)",
                        zIndex: 9999,
                    }}
                >
                    {" "}
                    <Spinner color="primary" />
                </div>
            )}
            <div className="header gradient-dark pb-6 pt-5 pt-md-8">
                <Container fluid>
                    <div className="header-body"></div>
                </Container>
            </div>
            <Container className="mt--9" fluid>
                <Card className="shadow">
                    <CardHeader className="border-0">
                        <h3 className="mb-0">{currentView}</h3>
                        <div className="col text-right">
                            {/* <Button
                                color="primary"
                                onClick={() => setModalOpen(true)}
                            >
                                Upload CSV
                            </Button> */}

                            <UncontrolledDropdown>
                                <DropdownToggle
                                    caret
                                    color="secondary"
                                    id="dropdownMenuButton"
                                    type="button"
                                >
                                    View
                                </DropdownToggle>

                                <DropdownMenu aria-labelledby="dropdownMenuButton">
                                    <DropdownItem onClick={() => setCurrentView('Bordereaux Reporting')}>
                                        Bordereaux Reporting
                                    </DropdownItem>

                                    <DropdownItem onClick={() => setCurrentView('Vessel Database')}>
                                        Vessel Database
                                    </DropdownItem>

                                    <DropdownItem onClick={() => setCurrentView('Signing')}>
                                        Signing
                                    </DropdownItem>
                                </DropdownMenu>
                            </UncontrolledDropdown>

                        </div>
                    </CardHeader>
                    <CardBody>
                        <Table className="align-items-center table-flush">
                            {renderTableContent()}
                        </Table>
                    </CardBody>

                    {/* <CardBody>
                        <Table className="align-items-center table-flush">
                            <thead className="thead-light">
                                <tr>
                                    <th scope="col">Reporting Period</th>
                                    <th scope="col">Business Unit</th>
                                    <th scope="col">Upload Date</th>
                                    <th scope="col">Status</th>
                                    <th scope="col" className="text-end">
                                        <Button
                                            size="sm"
                                            color="primary"
                                            onClick={() => setModalOpen(true)}
                                        >
                                            Upload CSV
                                        </Button>

                                    </th>
                                </tr>
                            </thead>
                            <tbody>
                                {Array.isArray(csvFiles) && csvFiles.map((file, index) => (
                                    <tr key={index}>
                                        <td>{file.reportingPeriod}</td>
                                        <td>{file.businessUnit}</td>
                                        <td>{new Date(file.uploadedAt).toLocaleString()}</td>
                                        <td>
                                            <Badge color={file.status === 'Processed' ? 'success' : 'warning'}>
                                                {file.status}
                                            </Badge>
                                        </td>
                                        <td className="px-6 py-4 whitespace-nowrap">
                                       
                                            <UncontrolledDropdown>
                                                <DropdownToggle
                                                    className="btn-icon-only text-light"
                                                    role="button"
                                                    size="sm"
                                                    color=""
                                                    onClick={(e) => e.preventDefault()}
                                                >
                                                    <i className="fas fa-ellipsis-v" />
                                                </DropdownToggle>
                                                <DropdownMenu className="dropdown-menu-arrow" right>
                                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "claimsummary.csv") && file._files.some(f => f.fileName === "premiumsummary.csv") && (
                                                        <DropdownItem onClick={() => {

                                                            getSummary(file._id, "statement_of_accounts.csv");
                                                        }}>
                                                            Generate Statement
                                                        </DropdownItem>
                                                    )}
                                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "claimsummary.csv") && (
                                                        <>
                                                            <DropdownItem onClick={() => {
                                                                getSummary(file._id, "claimsummary.csv");
                                                            }}>
                                                                View Claim Summary
                                                            </DropdownItem>
                                                        </>
                                                    )}

                                                    {file && Array.isArray(file._files) && file._files.some(f => f.fileName === "premiumsummary.csv") && (
                                                        <>
                                                            <DropdownItem onClick={() => {
                                                                getSummary(file._id, "premiumsummary.csv");
                                                            }}>
                                                                View Premium Summary
                                                            </DropdownItem>
                                                          </>
                                                    )}


                                                </DropdownMenu>


                                            </UncontrolledDropdown>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </Table>
                    </CardBody> */}


                </Card>
            </Container>

            <Modal isOpen={modalOpen} toggle={() => setModalOpen(!modalOpen)}>
                <Form onSubmit={handleFileUpload}>
                    <ModalHeader toggle={() => setModalOpen(!modalOpen)}>Upload CSV File</ModalHeader>
                    <ModalBody>
                        <FormGroup>
                            {
                                !file ? (
                                    <> <div className=" custom-file">
                                        <input
                                            className=" custom-file-input"
                                            id="customFileLang"
                                            lang="en"
                                            type="file"
                                            accept=".xlsx"
                                            // onChange={(e) => setFile(e.target.files[0])}
                                            onChange={handleFileChange}
                                        ></input>

                                        <label
                                            className=" custom-file-label"
                                            htmlFor="customFileLang"
                                        >
                                            Select file
                                        </label>
                                    </div></>
                                ) : (
                                    //button to remove file
                                    <div>
                                        <span>{file.name}</span>
                                        <button
                                            type="button"
                                            onClick={() => setFile(null)}
                                        >
                                            Remove
                                        </button>
                                    </div>
                                )



                            }
                        </FormGroup>
                        <FormGroup>
                            <Label for="reportType">Bordereaux Type</Label>
                            <Input
                                type="select"
                                name="reportType"
                                id="reportType"
                                value={reportType}
                                onChange={(e) => setReportType(e.target.value)}
                                required
                            >
                                <option value="">Select a bordereaux type</option>
                                <option value="premiums_paid">Premiums Paid Bordereaux</option>
                                <option value="claims_paid">Claims Paid Bordereaux</option>
                            </Input>
                        </FormGroup>
                        <FormGroup>
                            <Label for="reportingPeriod">Reporting Period</Label>
                            <Input
                                type="month"
                                name="reportingPeriod"
                                id="reportingPeriod"
                                value={reportingPeriod}
                                onChange={(e) => setReportingPeriod(e.target.value)}
                                required
                            />
                        </FormGroup>
                        <FormGroup>
                            <Label for="businessUnit">Business Unit</Label>

                            <Input
                                type="select"
                                name="businessUnit"
                                id="businessUnit"
                                value={businessUnit}
                                onChange={(e) => setBusinessUnit(e.target.value)}
                                required
                            >
                                <option value="" disabled>Select a business unit</option>
                                <option>Business Unit 1</option>
                                <option>Business Unit 2</option>
                                <option>Business Unit 3</option>
                            </Input>
                        </FormGroup>
                    </ModalBody>
                    <Alert color="info" isOpen={errorAlert}>
                        <span className="alert-icon">
                            <i className="ni ni-like-2"></i>
                        </span>
                        <span className="alert-text">
                            <strong>Default!</strong>{" "}
                            This is a default alert—check it out!
                        </span>
                        <button
                            type="button"
                            className="close"
                            data-dismiss="alert"
                            aria-label="Close"
                            onClick={() => { setErrorAlert(false) }}
                        >
                            <span aria-hidden="true">&times;</span>
                        </button>
                    </Alert>
                    <ModalFooter>
                        <Button color="secondary" onClick={() => setModalOpen(false)}>Cancel</Button>
                        <Button color="primary" type="submit">Upload</Button>
                    </ModalFooter>
                </Form>
            </Modal>
            <Modal isOpen={csvViewModalOpen} toggle={() => setCsvViewModalOpen(false)} size="xl" style={{ maxWidth: '100%', margin: '0', height: '100vh' }}>
                <ModalHeader toggle={() => setCsvViewModalOpen(false)}>CSV Content</ModalHeader>
                <ModalBody style={{ overflowY: 'auto', maxHeight: 'calc(100vh - 120px)' }}>
                    {csvData && (
                        <Table responsive>
                            <thead>
                                <tr>
                                    {csvData.headers.map((header, index) => (
                                        <th key={index}>{header}</th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody>
                                {csvData.data.map((row, rowIndex) => (
                                    <tr key={rowIndex}>
                                        {csvData.headers.map((header, cellIndex) => (
                                            <td key={cellIndex}>{row[header]}</td>
                                        ))}
                                    </tr>
                                ))}
                            </tbody>
                        </Table>
                    )}
                </ModalBody>
            </Modal>
            {/* Modal for viewing screening info */}
            <Modal isOpen={screenModal} toggle={() => setScreenModal(false)} style={{ maxWidth: "1000px", width: "100%" }}>
                <ModalHeader toggle={() => setScreenModal(false)}>Screening Report</ModalHeader>
                <ModalBody>
                    <Row>
                        <Col lg="12">
                            <FormGroup>
                                <label className="form-control-label">
                                    Potential Matches
                                </label>
                                <Input
                                    className="form-control-alternative"
                                    disabled
                                    type="textarea"
                                    value={(() => {
                                        if (!selectedVessel?.OFAC) return "N/A";

                                        const ofacData = selectedVessel.OFAC;
                                        let displayText = "Sources Used:\n";

                                        // Add sources information with download date
                                        if (ofacData.sources || ofacData.sourcesUsed) {
                                            const sources = ofacData.sources || ofacData.sourcesUsed;
                                            sources.forEach(source => {
                                                displayText += `- ${source.source}${source.name ? ` (${source.name})` : ''}\n`;
                                                displayText += `  Published: ${source.publishDate}\n`;
                                                if (source.downloadDate) {
                                                    displayText += `  Downloaded: ${source.downloadDate}\n`;
                                                }
                                                displayText += '\n';
                                            });
                                        }

                                        displayText += "Potential Matches:\n";

                                        // Handle matches with new structure
                                        if (ofacData.results) {
                                            const results = ofacData.results[0];
                                            const matchCount = results?.matchCount?.$numberInt || 0;

                                            if (matchCount === "0" || !results.matches) {
                                                displayText += "No potential matches found.\n";
                                            } else {
                                                results.matches.forEach((match, index) => {
                                                    displayText += `\nMatch ${index + 1}:\n`;
                                                    if (match.name) displayText += `Name: ${match.name}\n`;
                                                    if (match.type) displayText += `Type: ${match.type}\n`;
                                                    if (match.programs) displayText += `Programs: ${match.programs.join(', ')}\n`;
                                                    if (match.remarks) displayText += `Remarks: ${match.remarks}\n`;
                                                    if (match.vesselDetails) {
                                                        const vessel = match.vesselDetails;
                                                        if (vessel.flag) displayText += `Flag: ${vessel.flag}\n`;
                                                        if (vessel.imoNumber) displayText += `IMO: ${vessel.imoNumber}\n`;
                                                    }
                                                    if (match.categories) displayText += `Categories: ${match.categories.join(', ')}\n`;
                                                    if (match.source) displayText += `Source: ${match.source}\n`;
                                                    if (match.identifications) {
                                                        match.identifications.forEach(id => {
                                                            displayText += `${id.type} ${id.idNumber}\n`;
                                                        });
                                                    }
                                                });
                                            }
                                        }

                                        return displayText;
                                    })()}
                                    rows="10"
                                    style={{ whiteSpace: 'pre-wrap' }}
                                />
                            </FormGroup>
                        </Col>
                        <Col className="text-center">
                            <Button
                                color="primary"
                                disabled={selectedVessel?.OFAC}
                                className={selectedVessel?.OFAC ? 'opacity-50 cursor-not-allowed' : ''}
                                onClick={() => { screenVessel(selectedVessel) }}
                            >  Start Screening
                            </Button>
                        </Col>
                    </Row>
                </ModalBody>
            </Modal>
            <Modal>
                <ModalHeader toggle={() => setTrackModal(false)}>Tracking Report</ModalHeader>
                <ModalBody> <div className="text-center">
                    <Row className="mt-2 mb-2">
                        {/* {positions.length > 0 ? <Map positions={positions} /> : <p>Loading...</p>} */}
                        {positions.map((position, index) => (
                            <Col md="4" key={index}
                                style={{ marginBottom: '20px' }}> {/* Adjust marginBottom as needed */}
                                <Card>
                                    <CardHeader>Position {index + 1}</CardHeader> {/* Add a CardHeader for the title */}
                                    <CardBody>
                                        <CardText>Latitude: {position.lat}</CardText>
                                        <CardText>Longitude: {position.lon}</CardText>
                                        <CardText>Speed: {position.speed}</CardText>
                                        <CardText>Course: {position.course}</CardText>
                                        <CardText>Heading: {position.heading}</CardText>
                                        <CardText>Destination: {position.destination}</CardText>
                                        <CardText>Last Position
                                            Epoch: {position.last_position_epoch}</CardText>
                                        <CardText>Last Position
                                            UTC: {position.last_position_UTC}</CardText>
                                        <CardText>
                                            <a
                                                href={`https://www.google.com/maps?q=${position.lat},${position.lon}`}
                                                target="_blank"
                                                rel="noopener noreferrer"
                                            >
                                                View on Google Maps
                                            </a>
                                        </CardText>
                                    </CardBody>
                                </Card>
                            </Col>
                        ))}
                    </Row>
                </div></ModalBody>
            </Modal>

        </Admin>
    );
}

export default CSV;
