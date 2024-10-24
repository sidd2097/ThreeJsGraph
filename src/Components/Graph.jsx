import React, { useEffect, useRef } from 'react';
import * as THREE from 'three';
import { read, utils } from 'xlsx';

const Graph = () => {
  const mountRef = useRef(null);

  useEffect(() => {
    const loadExcelData = async () => {
      try {
        const response = await fetch('/data/ThreeGraph.xlsx');
        const arrayBuffer = await response.arrayBuffer();
        const workbook = read(arrayBuffer, { type: 'array' });

        // Parse Nodes Sheet
        const nodesSheet = workbook.Sheets['Nodes'];
        if (!nodesSheet) throw new Error("Nodes sheet not found in the Excel file");
        const nodesData = utils.sheet_to_json(nodesSheet);

        // Parse Members Sheet
        const membersSheet = workbook.Sheets['Members'];
        if (!membersSheet) throw new Error("Members sheet not found in the Excel file");
        const membersData = utils.sheet_to_json(membersSheet);

        // Initialize Three.js scene
        const scene = new THREE.Scene();
        const camera = new THREE.PerspectiveCamera(75, window.innerWidth / window.innerHeight, 0.1, 1000);
        const renderer = new THREE.WebGLRenderer();
        renderer.setSize(window.innerWidth, window.innerHeight);
        mountRef.current.appendChild(renderer.domElement);

        // Add light to the scene
        const light = new THREE.PointLight(0xffffff, 1);
        light.position.set(10, 10, 10);
        scene.add(light);

        // Create a material for the nodes
        const nodeMaterial = new THREE.MeshBasicMaterial({ color: 0x007bff });
        const geometry = new THREE.SphereGeometry(0.5, 16, 16);

        // Create node objects and add them to the scene
        const nodes = {};
        nodesData.forEach(node => {
          const { Node, X, Y, Z } = node;
          const mesh = new THREE.Mesh(geometry, nodeMaterial);
          mesh.position.set(X, Y, Z);
          scene.add(mesh);
          nodes[Node] = mesh.position;  // Store the position by node ID
        });

        // Create line material for members
        const lineMaterial = new THREE.LineBasicMaterial({ color: 0xff0000 });

        // Add members as lines connecting nodes
        membersData.forEach(member => {
          const { start, end } = member;
          const points = [];
          if (nodes[start] && nodes[end]) {
            points.push(new THREE.Vector3(nodes[start].x, nodes[start].y, nodes[start].z));
            points.push(new THREE.Vector3(nodes[end].x, nodes[end].y, nodes[end].z));

            const memberGeometry = new THREE.BufferGeometry().setFromPoints(points);
            const line = new THREE.Line(memberGeometry, lineMaterial);
            scene.add(line);
          }
        });

        camera.position.z = 100;

        // Animation loop
        const animate = () => {
          requestAnimationFrame(animate);
          renderer.render(scene, camera);
        };

        animate();
      } catch (error) {
        console.error("Error loading Excel data:", error);
      }
    };

    loadExcelData();

    return () => {
      if (mountRef.current && mountRef.current.firstChild) {
        mountRef.current.removeChild(mountRef.current.firstChild);
      }
    };
  }, []);

  return <div ref={mountRef}></div>;
};

export default Graph;
